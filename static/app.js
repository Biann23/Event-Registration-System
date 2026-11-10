/* LSPU Event Management System: progressive enhancement. Every page still works without it. */
(function () {
  "use strict";
  var doc = document;
  function all(sel, root) { return Array.prototype.slice.call((root || doc).querySelectorAll(sel)); }
  doc.documentElement.classList.add("js");

  /* Dismissible flash messages */
  all(".flash-close").forEach(function (b) {
    b.addEventListener("click", function () { b.closest(".flash").remove(); });
  });

  /* Dropdown menus (<details>): close on outside click or Escape */
  var menus = all("details.user-menu, details.more");
  doc.addEventListener("click", function (e) {
    menus.forEach(function (m) { if (m.open && !m.contains(e.target)) m.open = false; });
  });
  doc.addEventListener("keydown", function (e) {
    if (e.key !== "Escape") return;
    menus.forEach(function (m) { if (m.open) { m.open = false; var s = m.querySelector("summary"); if (s) s.focus(); } });
  });

  /* Show / hide password */
  all('input[type="password"]').forEach(function (input) {
    var wrap = doc.createElement("span");
    wrap.className = "pw";
    input.parentNode.insertBefore(wrap, input);
    wrap.appendChild(input);
    var btn = doc.createElement("button");
    btn.type = "button"; btn.className = "pw-toggle"; btn.setAttribute("aria-label", "Show password");
    btn.innerHTML = '<svg class="ico" viewBox="0 0 24 24" aria-hidden="true"><path d="M2 12s3.6-7 10-7 10 7 10 7-3.6 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/></svg>';
    btn.addEventListener("click", function () {
      var show = input.type === "password";
      input.type = show ? "text" : "password";
      btn.setAttribute("aria-label", show ? "Hide password" : "Show password");
      btn.setAttribute("aria-pressed", show ? "true" : "false");
    });
    wrap.appendChild(btn);
  });

  /* Event status chips on the student board */
  var chips = doc.querySelector("[data-chips]");
  if (chips) {
    var cards = all("[data-status]");
    var none = doc.getElementById("chip-empty");
    all("button", chips).forEach(function (btn) {
      btn.addEventListener("click", function () {
        all("button", chips).forEach(function (b) { b.setAttribute("aria-pressed", b === btn ? "true" : "false"); });
        var want = btn.getAttribute("data-value"), shown = 0;
        cards.forEach(function (c) {
          var st = c.getAttribute("data-status");
          var hit = want === "all" ||
            (want === "now" && (st === "ongoing" || st === "lunch-break")) ||
            (want === "soon" && st === "not-started-yet") ||
            (want === "ended" && (st === "ended" || st === "archived"));
          c.hidden = !hit; if (hit) shown++;
        });
        if (none) none.hidden = shown !== 0;
      });
    });
  }

  /* Live filter for long tables: <input data-filter="#table, .other-table"> */
  all("[data-filter]").forEach(function (input) {
    var targets = all(input.getAttribute("data-filter"));
    var out = doc.getElementById(input.getAttribute("data-count") || "none");
    if (!targets.length) return;
    input.addEventListener("input", function () {
      var q = input.value.trim().toLowerCase(), shown = 0;
      targets.forEach(function (target) {
        var visible = 0;
        all("tbody tr:not(.group-row)", target).forEach(function (tr) {
          var hit = !q || tr.textContent.toLowerCase().indexOf(q) !== -1;
          tr.hidden = !hit; if (hit) { visible++; shown++; }
        });
        var block = target.closest(".group-block");
        if (block) { block.hidden = !!q && visible === 0; if (q && visible) block.open = true; }
      });
      if (out) out.textContent = q ? shown + " match" + (shown === 1 ? "" : "es") : "";
    });
  });

  /* Confirmation dialog for destructive actions: <form data-confirm="..."> */
  var dlg = doc.getElementById("confirm-dialog");
  all("form[data-confirm]").forEach(function (form) {
    form.addEventListener("submit", function (e) {
      if (form.getAttribute("data-confirmed") === "1") return;
      e.preventDefault();
      var text = form.getAttribute("data-confirm");
      if (!dlg || typeof dlg.showModal !== "function") {
        if (window.confirm(text)) { form.setAttribute("data-confirmed", "1"); form.submit(); }
        return;
      }
      dlg.querySelector("[data-d-title]").textContent = form.getAttribute("data-confirm-title") || "Please confirm";
      dlg.querySelector("[data-d-text]").textContent = text;
      dlg.querySelector("[data-d-ok]").textContent = form.getAttribute("data-confirm-ok") || "Confirm";
      dlg.returnValue = "";
      dlg.onclose = function () {
        if (dlg.returnValue === "ok") {
          form.setAttribute("data-confirmed", "1");
          if (form.requestSubmit) form.requestSubmit(); else form.submit();
        }
      };
      dlg.showModal();
    });
  });

  /* Prevent double submits: <form data-lock> */
  all("form[data-lock]").forEach(function (form) {
    form.addEventListener("submit", function (e) {
      if (e.defaultPrevented) return;
      var btn = form.querySelector('button[type="submit"]');
      if (!btn) return;
      if (btn.disabled) { e.preventDefault(); return; }
      var label = btn.getAttribute("data-busy") || "Working";
      setTimeout(function () {
        btn.disabled = true; btn.setAttribute("aria-busy", "true");
        btn.innerHTML = '<span class="spinner" aria-hidden="true"></span> ' + label;
      }, 0);
    });
  });

  /* Print buttons */
  all("[data-print]").forEach(function (b) { b.addEventListener("click", function () { window.print(); }); });
})();
