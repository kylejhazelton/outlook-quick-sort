/*
 * Quick Sort: one-click categories for Outlook.
 *
 * Uses only Office.js (Mailbox 1.8), so no sign-in or app registration is needed.
 * Opened outside Outlook (e.g. straight in a browser), it runs in preview mode
 * with sample data so you can check the layout.
 */
(function () {
  "use strict";

  var TOAST_MS = 6000;
  // ---- Top 5 tuning ----
  // Each time you add a category, every score is multiplied by DECAY and the
  // chosen one gets +1. A pick is worth half as much after ~14 later picks.
  var TOP_N = 5;
  var DECAY = 0.95;
  // A newcomer must beat the weakest Top 5 member by this much to replace it,
  // so two close categories don't keep swapping places.
  var MARGIN = 0.5;
  var MIN_SCORE = 0.01; // scores below this are forgotten
  var TOP_KEY = "topCategories";
  var OLD_RECENT_KEY = "recentCategories"; // from the earlier "Recent" version

  // Outlook's category colour presets.
  var COLORS = {
    None: "#8a8a8a",
    Preset0: "#e74856", Preset1: "#ff8c00", Preset2: "#ab7b5a", Preset3: "#fff100",
    Preset4: "#47d041", Preset5: "#30c6cc", Preset6: "#73aa24", Preset7: "#4cafe8",
    Preset8: "#8764b8", Preset9: "#f472d0", Preset10: "#c0c7d0", Preset11: "#9fb2c7",
    Preset12: "#a0a0a0", Preset13: "#626262", Preset14: "#1f1f1f", Preset15: "#a4262c",
    Preset16: "#ca5010", Preset17: "#8e562e", Preset18: "#c19c00", Preset19: "#0b6a0b",
    Preset20: "#038387", Preset21: "#5b7e1d", Preset22: "#004e8c", Preset23: "#5c2e91",
    Preset24: "#bf0077"
  };

  var state = {
    mock: false,
    masterCategories: [],
    itemCategories: new Set(),
    scores: {}, // category name -> weighted score
    top: []     // names currently in the Top 5
  };

  function compareNames(a, b) {
    // Alphabetical, ignoring case; numbers sort naturally ("2" before "10").
    return a.localeCompare(b, undefined, { sensitivity: "base", numeric: true });
  }

  function $(id) { return document.getElementById(id); }

  function el(tag, cls, text) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text != null) n.textContent = text;
    return n;
  }

  function delay(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  function officeAsync(call) {
    return new Promise(function (resolve, reject) {
      call(function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) resolve(r.value);
        else reject(new Error((r.error && r.error.message) || "Outlook request failed"));
      });
    });
  }

  // ---------- mailbox adapters (real Outlook vs. preview) ----------

  var real = {
    hasItem: function () { return !!Office.context.mailbox.item; },
    subject: function () {
      var item = Office.context.mailbox.item;
      return item ? (item.subject || "(no subject)") : null;
    },
    masterCategories: function () {
      return officeAsync(function (cb) { Office.context.mailbox.masterCategories.getAsync(cb); });
    },
    itemCategories: function () {
      return officeAsync(function (cb) { Office.context.mailbox.item.categories.getAsync(cb); });
    },
    addCategory: function (name) {
      return officeAsync(function (cb) { Office.context.mailbox.item.categories.addAsync([name], cb); });
    },
    removeCategory: function (name) {
      return officeAsync(function (cb) { Office.context.mailbox.item.categories.removeAsync([name], cb); });
    },
    // Roaming settings are stored in your mailbox, so Top 5 follows you
    // between new Outlook on desktop and Outlook on the web.
    loadTop: function () {
      var rs = Office.context.roamingSettings;
      var v = rs.get(TOP_KEY);
      if (v && typeof v === "object" && v.scores) return v;
      // First run after the upgrade: seed scores from the old Recent list
      // (small values, so real usage quickly takes over).
      var old = rs.get(OLD_RECENT_KEY);
      var scores = {};
      if (Array.isArray(old)) {
        old.slice(0, TOP_N).forEach(function (n, i) { scores[n] = Math.pow(DECAY, i); });
      }
      return { scores: scores, top: [] };
    },
    saveTop: function (data) {
      var rs = Office.context.roamingSettings;
      rs.set(TOP_KEY, data);
      rs.remove(OLD_RECENT_KEY);
      rs.saveAsync();
    }
  };

  var mockItemCats = new Set(["Follow up"]);
  var mock = {
    hasItem: function () { return true; },
    subject: function () { return "Q3 budget review: notes from Tuesday"; },
    masterCategories: async function () {
      return [
        { displayName: "Follow up", color: "Preset0" },
        { displayName: "Waiting on others", color: "Preset1" },
        { displayName: "Finance", color: "Preset4" },
        { displayName: "Travel", color: "Preset7" },
        { displayName: "Personal", color: "Preset8" },
        { displayName: "Reference", color: "Preset12" }
      ];
    },
    itemCategories: async function () {
      return Array.from(mockItemCats).map(function (n) { return { displayName: n }; });
    },
    addCategory: async function (n) { await delay(120); mockItemCats.add(n); },
    removeCategory: async function (n) { await delay(120); mockItemCats.delete(n); },
    loadTop: function () {
      return { scores: { "Follow up": 3.2, "Travel": 2.1, "Finance": 1.4 }, top: [] };
    },
    saveTop: function () { /* preview: not persisted */ }
  };

  var api = real;

  // ---------- rendering ----------

  function renderItem() {
    var subject = api.subject();
    $("subject").textContent = subject || "No message selected";
    $("subject").title = subject || "";
  }

  function categoryButton(c, hasItem) {
    var on = state.itemCategories.has(c.displayName);
    var b = el("button", "chip");
    b.type = "button";
    b.setAttribute("aria-pressed", on ? "true" : "false");
    b.title = (on ? "Remove " : "Add ") + c.displayName;
    b.disabled = !hasItem;
    var dot = el("span", "dot");
    dot.style.setProperty("--c", COLORS[c.color] || COLORS.None);
    b.appendChild(dot);
    b.appendChild(el("span", "label", c.displayName));
    b.appendChild(el("span", "check", on ? "✓" : ""));
    b.addEventListener("click", function () { toggleCategory(c.displayName); });
    return b;
  }

  function renderCategories() {
    var hasItem = api.hasItem();

    // Top 5: shown alphabetically so positions stay put as scores shift.
    var byName = {};
    state.masterCategories.forEach(function (c) { byName[c.displayName] = c; });
    var top = state.top
      .filter(function (n) { return byName[n]; })
      .sort(compareNames)
      .map(function (n) { return byName[n]; });
    var topWrap = $("recent");
    topWrap.textContent = "";
    top.forEach(function (c) { topWrap.appendChild(categoryButton(c, hasItem)); });
    $("recent-section").hidden = top.length === 0;

    // All: every category, alphabetical.
    var wrap = $("categories");
    wrap.textContent = "";
    state.masterCategories.forEach(function (c) { wrap.appendChild(categoryButton(c, hasItem)); });
    $("cat-empty").hidden = state.masterCategories.length > 0;
  }

  // Works out the Top 5 from the scores. Existing members keep their spot
  // unless a newcomer beats the weakest member by MARGIN.
  // Returns true if the membership changed.
  function recomputeTop() {
    var exists = {};
    state.masterCategories.forEach(function (c) { exists[c.displayName] = true; });
    var score = function (n) { return state.scores[n] || 0; };

    var members = state.top.filter(function (n) { return exists[n] && score(n) > 0; });
    var others = Object.keys(state.scores)
      .filter(function (n) { return exists[n] && members.indexOf(n) === -1 && score(n) > 0; })
      .sort(function (a, b) { return score(b) - score(a); });

    // Fill empty slots with the highest scorers.
    while (members.length < TOP_N && others.length) members.push(others.shift());

    // Let a clearly stronger newcomer replace the weakest member.
    while (others.length) {
      var weakest = members.reduce(function (w, n) { return score(n) < score(w) ? n : w; }, members[0]);
      if (score(others[0]) > score(weakest) + MARGIN) {
        members[members.indexOf(weakest)] = others.shift();
        others.push(weakest);
        others.sort(function (a, b) { return score(b) - score(a); });
      } else {
        break;
      }
    }

    var changed = members.slice().sort().join("\n") !== state.top.slice().sort().join("\n");
    state.top = members;
    return changed;
  }

  function saveTop() {
    try { api.saveTop({ scores: state.scores, top: state.top }); } catch (e) { /* not critical */ }
  }

  function recordPick(name) {
    var next = {};
    Object.keys(state.scores).forEach(function (n) {
      var s = state.scores[n] * DECAY;
      if (s >= MIN_SCORE) next[n] = Math.round(s * 1000) / 1000;
    });
    next[name] = Math.round(((next[name] || 0) + 1) * 1000) / 1000;
    state.scores = next;
    recomputeTop();
    saveTop();
  }

  var toastTimer = null;
  function toast(msg) {
    $("toast-msg").textContent = msg;
    $("toast").hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { $("toast").hidden = true; }, TOAST_MS);
  }

  // ---------- actions ----------

  async function loadCategories() {
    try {
      state.masterCategories = ((await api.masterCategories()) || []).slice().sort(function (a, b) {
        return compareNames(a.displayName, b.displayName);
      });
      // Fill free slots / drop deleted categories from the Top 5.
      if (recomputeTop()) saveTop();
    } catch (e) {
      state.masterCategories = [];
      toast("Couldn't load categories: " + e.message);
    }
    await loadItemCategories();
  }

  async function loadItemCategories() {
    state.itemCategories = new Set();
    if (api.hasItem()) {
      try {
        var cats = (await api.itemCategories()) || [];
        cats.forEach(function (c) { state.itemCategories.add(c.displayName); });
      } catch (e) { /* item may have gone away */ }
    }
    renderCategories();
  }

  async function toggleCategory(name) {
    if (!api.hasItem()) return;
    var had = state.itemCategories.has(name);
    // Optimistic update so the click feels instant; rolled back on failure.
    if (had) state.itemCategories.delete(name); else state.itemCategories.add(name);
    renderCategories();
    try {
      if (had) {
        await api.removeCategory(name);
      } else {
        await api.addCategory(name);
        recordPick(name);
        renderCategories();
      }
    } catch (e) {
      if (had) state.itemCategories.add(name); else state.itemCategories.delete(name);
      renderCategories();
      toast("Couldn't update category: " + e.message);
    }
  }

  async function onItemChanged() {
    renderItem();
    await loadItemCategories();
  }

  // ---------- startup ----------

  async function start(inOutlook) {
    state.mock = !inOutlook;
    api = state.mock ? mock : real;
    $("mode").hidden = !state.mock;
    $("refresh").addEventListener("click", loadCategories);

    // Keep keyboard focus in Outlook's message list: cancelling mousedown on a
    // button stops the click from moving focus into this panel, while the click
    // itself still fires. (Tab-key navigation into the panel still works.)
    document.addEventListener("mousedown", function (e) {
      if (e.button === 0 && e.target.closest && e.target.closest("button")) e.preventDefault();
    }, true);
    try {
      var saved = api.loadTop();
      state.scores = saved.scores || {};
      state.top = Array.isArray(saved.top) ? saved.top : [];
    } catch (e) { state.scores = {}; state.top = []; }

    renderItem();
    await loadCategories();

    if (!state.mock) {
      // Fires when you select a different email while the pane is pinned.
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged);
    }
  }

  if (window.Office && Office.onReady) {
    Office.onReady(function (info) {
      start(info.host === Office.HostType.Outlook && !!(Office.context && Office.context.mailbox));
    });
  } else {
    // office.js couldn't load (offline?) - still show the preview.
    start(false);
  }
})();
