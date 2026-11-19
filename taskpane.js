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
    itemCategories: new Set()
  };

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
    removeCategory: async function (n) { await delay(120); mockItemCats.delete(n); }
  };

  var api = real;

  // ---------- rendering ----------

  function renderItem() {
    var subject = api.subject();
    $("subject").textContent = subject || "No message selected";
    $("subject").title = subject || "";
  }

  function renderCategories() {
    var wrap = $("categories");
    wrap.textContent = "";
    var hasItem = api.hasItem();
    state.masterCategories.forEach(function (c) {
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
      wrap.appendChild(b);
    });
    $("cat-empty").hidden = state.masterCategories.length > 0;
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
        // Alphabetical, ignoring case; numbers sort naturally ("2" before "10").
        return a.displayName.localeCompare(b.displayName, undefined, { sensitivity: "base", numeric: true });
      });
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
      if (had) await api.removeCategory(name); else await api.addCategory(name);
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
