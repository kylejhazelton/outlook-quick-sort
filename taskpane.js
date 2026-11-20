/*
 * Quick Sort: one-click categories for Outlook.
 *
 * Uses only Office.js, so no sign-in or app registration is needed.
 *  - One email selected: Mailbox 1.8 category APIs on the open item.
 *  - Several selected:   Mailbox 1.15 loadItemByIdAsync, one email at a time
 *                        (Outlook only allows one loaded item at once).
 *
 * Opened outside Outlook (e.g. straight in a browser), it runs in preview mode
 * with sample data. Add ?multi=4 to the URL to preview a 4-email selection.
 */
(function () {
  "use strict";

  var TOAST_MS = 6000;
  // ---- Top 5 tuning ----
  // Each time you add a category, every score is multiplied by DECAY and the
  // chosen one gets +1 per email it was applied to (+12 for 12 selected emails).
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
    itemCategories: new Set(), // single-email mode
    // Multi-email mode (null when one email is selected):
    //   { items: [{id, subject}], cats: {id: Set}, ready: bool }
    multi: null,
    busy: false,               // applying to several emails
    status: "",                // progress text under the header
    scores: {},                // category name -> weighted score
    top: []                    // names currently in the Top 5
  };
  var selectionGen = 0; // bumps on every selection change, to drop stale work

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

  function names(cats) {
    return new Set((cats || []).map(function (c) { return c.displayName; }));
  }

  // Outlook allows only one loaded item at a time, so every loaded-item
  // operation goes through this queue.
  var loadQueue = Promise.resolve();
  function queued(task) {
    var run = loadQueue.then(task);
    loadQueue = run.catch(function () {});
    return run;
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

    supportsMulti: function () {
      return Office.context.requirements.isSetSupported("Mailbox", "1.15");
    },
    selectedItems: function () {
      return officeAsync(function (cb) { Office.context.mailbox.getSelectedItemsAsync(cb); })
        .then(function (list) {
          return (list || []).map(function (s) { return { id: s.itemId, subject: s.subject }; });
        });
    },
    // Loads one selected email, runs fn(ops), then unloads it.
    withItem: function (id, fn) {
      return queued(async function () {
        var item = await officeAsync(function (cb) { Office.context.mailbox.loadItemByIdAsync(id, cb); });
        try {
          return await fn({
            categories: function () { return officeAsync(function (cb) { item.categories.getAsync(cb); }); },
            add: function (n) { return officeAsync(function (cb) { item.categories.addAsync([n], cb); }); },
            remove: function (n) { return officeAsync(function (cb) { item.categories.removeAsync([n], cb); }); }
          });
        } finally {
          await new Promise(function (r) { item.unloadAsync(function () { r(); }); });
        }
      });
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

  // ---- preview mode ----
  var mockMultiCount = (function () {
    var m = /[?&]multi=(\d+)/.exec(location.search);
    return m ? Math.min(100, parseInt(m[1], 10)) : 0;
  })();
  var mockStore = {}; // id -> Set of categories
  var mockSubjects = ["Q3 budget review: notes from Tuesday", "Site visit photos", "Invoice #4471",
    "Re: Lunch Friday?", "Updated drawings", "Travel itinerary"];
  for (var i = 0; i < Math.max(1, mockMultiCount); i++) {
    mockStore["mock-" + i] = new Set(i % 2 === 0 ? ["Follow up"] : ["Follow up", "Finance"]);
  }
  var mock = {
    hasItem: function () { return mockMultiCount < 2; },
    subject: function () { return mockMultiCount < 2 ? mockSubjects[0] : null; },
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
      return Array.from(mockStore["mock-0"]).map(function (n) { return { displayName: n }; });
    },
    addCategory: async function (n) { await delay(120); mockStore["mock-0"].add(n); },
    removeCategory: async function (n) { await delay(120); mockStore["mock-0"].delete(n); },
    supportsMulti: function () { return true; },
    selectedItems: async function () {
      return Object.keys(mockStore).map(function (id, i) {
        return { id: id, subject: mockSubjects[i % mockSubjects.length] };
      });
    },
    withItem: function (id, fn) {
      return queued(async function () {
        await delay(60);
        var set = mockStore[id];
        return fn({
          categories: async function () {
            return Array.from(set).map(function (n) { return { displayName: n }; });
          },
          add: async function (n) { await delay(60); set.add(n); },
          remove: async function (n) { await delay(60); set.delete(n); }
        });
      });
    },
    loadTop: function () {
      return { scores: { "Follow up": 3.2, "Travel": 2.1, "Finance": 1.4 }, top: [] };
    },
    saveTop: function () { /* preview: not persisted */ }
  };

  var api = real;

  // ---------- rendering ----------

  function renderHeader() {
    var text;
    if (state.multi) {
      var n = state.multi.items.length;
      text = n === 1 ? (state.multi.items[0].subject || "(no subject)") : n + " emails selected";
    } else {
      text = api.subject() || "No message selected";
    }
    $("subject").textContent = text;
    $("subject").title = text;
    $("status").textContent = state.status;
    $("status").hidden = !state.status;
  }

  // "on" = every selected email has it, "mixed" = some do, "off" = none.
  function categoryState(name) {
    if (!state.multi) return state.itemCategories.has(name) ? "on" : "off";
    var ids = Object.keys(state.multi.cats);
    var count = ids.filter(function (id) { return state.multi.cats[id].has(name); }).length;
    if (count === 0) return "off";
    // Compare against the emails read so far, so nothing looks "mixed" mid-read.
    return count === ids.length ? "on" : "mixed";
  }

  function canAct() {
    if (state.busy) return false;
    if (state.multi) return state.multi.ready;
    return api.hasItem();
  }

  function categoryButton(c, enabled) {
    var st = categoryState(c.displayName);
    var b = el("button", "chip");
    b.type = "button";
    b.setAttribute("aria-pressed", st === "on" ? "true" : st === "mixed" ? "mixed" : "false");
    b.title = (st === "on" ? "Remove " : "Add ") + c.displayName;
    b.disabled = !enabled;
    var dot = el("span", "dot");
    dot.style.setProperty("--c", COLORS[c.color] || COLORS.None);
    b.appendChild(dot);
    b.appendChild(el("span", "label", c.displayName));
    b.appendChild(el("span", "check", st === "on" ? "✓" : st === "mixed" ? "–" : ""));
    b.addEventListener("click", function () { toggleCategory(c.displayName); });
    return b;
  }

  function renderCategories() {
    renderHeader();
    var enabled = canAct();

    // Top 5: shown alphabetically so positions stay put as scores shift.
    var byName = {};
    state.masterCategories.forEach(function (c) { byName[c.displayName] = c; });
    var top = state.top
      .filter(function (n) { return byName[n]; })
      .sort(compareNames)
      .map(function (n) { return byName[n]; });
    var topWrap = $("recent");
    topWrap.textContent = "";
    top.forEach(function (c) { topWrap.appendChild(categoryButton(c, enabled)); });
    $("recent-section").hidden = top.length === 0;

    // All: every category, alphabetical.
    var wrap = $("categories");
    wrap.textContent = "";
    state.masterCategories.forEach(function (c) { wrap.appendChild(categoryButton(c, enabled)); });
    $("cat-empty").hidden = state.masterCategories.length > 0;
  }

  // ---------- Top 5 scoring ----------

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

  // One pick event: everything fades once, the chosen category gets +weight
  // (weight = number of selected emails).
  function recordPick(name, weight) {
    var next = {};
    Object.keys(state.scores).forEach(function (n) {
      var s = state.scores[n] * DECAY;
      if (s >= MIN_SCORE) next[n] = Math.round(s * 1000) / 1000;
    });
    next[name] = Math.round(((next[name] || 0) + (weight || 1)) * 1000) / 1000;
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

  // ---------- selection ----------

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
    renderCategories();
  }

  // Works out what's selected and loads its categories.
  async function refreshSelection() {
    var gen = ++selectionGen;
    var selected = null;
    if (api.supportsMulti()) {
      try { selected = await api.selectedItems(); } catch (e) { selected = null; }
      if (gen !== selectionGen) return;
    }

    var useMulti = selected && (selected.length > 1 || (selected.length === 1 && !api.hasItem()));
    if (!useMulti) {
      state.multi = null;
      state.status = "";
      state.itemCategories = new Set();
      if (api.hasItem()) {
        try { state.itemCategories = names(await api.itemCategories()); } catch (e) { /* item went away */ }
      }
      if (gen === selectionGen) renderCategories();
      return;
    }

    var m = { items: selected, cats: {}, ready: false };
    state.multi = m;
    for (var i = 0; i < selected.length; i++) {
      if (gen !== selectionGen) return; // selection changed: stop reading
      state.status = selected.length > 3 ? "Reading " + (i + 1) + " of " + selected.length + "…" : "";
      renderCategories();
      var id = selected[i].id;
      try {
        m.cats[id] = names(await api.withItem(id, function (ops) { return ops.categories(); }));
      } catch (e) {
        m.cats[id] = new Set();
      }
    }
    if (gen !== selectionGen) return;
    m.ready = true;
    state.status = "";
    renderCategories();
  }

  // ---------- actions ----------

  async function toggleCategory(name) {
    if (!canAct()) return;
    if (state.multi) return toggleMulti(name);

    var had = state.itemCategories.has(name);
    // Optimistic update so the click feels instant; rolled back on failure.
    if (had) state.itemCategories.delete(name); else state.itemCategories.add(name);
    renderCategories();
    try {
      if (had) {
        await api.removeCategory(name);
      } else {
        await api.addCategory(name);
        recordPick(name, 1);
        renderCategories();
      }
    } catch (e) {
      if (had) state.itemCategories.add(name); else state.itemCategories.delete(name);
      renderCategories();
      toast("Couldn't update category: " + e.message);
    }
  }

  // Several emails: if all have it, remove it from all; otherwise add it to
  // the ones that don't.
  async function toggleMulti(name) {
    var m = state.multi;
    var total = m.items.length;
    var removing = categoryState(name) === "on";
    var targets = m.items.filter(function (it) {
      var has = m.cats[it.id].has(name);
      return removing ? has : !has;
    });

    state.busy = true;
    var failed = 0;
    for (var i = 0; i < targets.length; i++) {
      if (state.multi === m) {
        state.status = (removing ? "Removing from " : "Tagging ") + (i + 1) + " of " + targets.length + "…";
        renderCategories();
      }
      var id = targets[i].id;
      try {
        await api.withItem(id, function (ops) { return removing ? ops.remove(name) : ops.add(name); });
        if (removing) m.cats[id].delete(name); else m.cats[id].add(name);
      } catch (e) {
        failed++;
      }
    }
    state.busy = false;
    if (!removing && failed < targets.length) recordPick(name, total);
    if (state.multi === m) state.status = "";
    renderCategories();

    var done = targets.length - failed;
    if (failed) {
      toast("Updated " + done + " of " + targets.length + " emails. " + failed + " couldn't be changed.");
    } else if (targets.length) {
      toast((removing ? "Removed " : "Added ") + name + (removing ? " from " : " to ") +
        done + " email" + (done === 1 ? "" : "s") + ".");
    }
  }

  // ---------- startup ----------

  async function start(inOutlook) {
    state.mock = !inOutlook;
    api = state.mock ? mock : real;
    $("mode").hidden = !state.mock;
    $("refresh").addEventListener("click", function () { loadCategories().then(refreshSelection); });

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

    await loadCategories();
    await refreshSelection();

    if (!state.mock) {
      // Fires when you select a different email while the pane is pinned.
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, refreshSelection);
      // Fires when you select several emails (Mailbox 1.13+).
      if (Office.context.requirements.isSetSupported("Mailbox", "1.13")) {
        Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, refreshSelection);
      }
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
