// Form-filling engine. Runs inside the claim website's page (injected by the side panel).
// A port of FormFiller in python-app/read_claims.py.
//
// The real site reloads the whole page after "Add claim", which destroys this script, so the
// extension cannot run the whole claim loop here. Instead the side panel keeps the loop and
// calls the small operations below one at a time (ClaimFillerEngine.run): "check", "state",
// "enter" (fill one claim and click Add claim) and "predetermine". Each call is stateless, so
// the panel can re-inject this file after a reload and carry on.
// fill() is the old single-page loop; it is kept for sites that do not reload (and for tests).
//
// Plain script (not a module) so it can be injected as-is.
(() => {
  if (globalThis.ClaimFillerEngine) return;

  const DEFAULTS = {
    fieldServiceDate: "Service date", // label reads "Service date (YYYY-MM-DD)"
    fieldService: "Service",
    fieldHours: "Number of hours",
    fieldCost: "Total cost", // label reads "Total cost ($)"
    buttonAdd: "Add claim",
    buttonPredetermine: "Predetermine",
    typingDelayMs: 30, // pause between keystrokes when typing hours/cost
    fieldTries: 3, // times one box is re-filled if it does not hold what was typed
    maxAttempts: 3, // tries per claim before giving up for good
    totalsWaitSeconds: 8, // how long to wait for the website total to update after Add claim
    pprMinPauseMs: 700, // minimum pause after changing the date/service (the site refreshes itself)
    pprQuietMs: 400, // the page must be unchanged this long before typing on
    pprMaxWaitMs: 8000, // never wait longer than this for a refresh to finish
  };

  // The real site shows only a bare running total ("$416.26") in the last row of the claims
  // table; "Total cost: $x" wording is accepted too. The last match on the page wins.
  const TOTAL_COST_PATTERNS = [/Total cost\s*:\s*\$\s*([\d,]*\.?\d+)/gi, /^[ \t]*\$\s*([\d,]*\.?\d+)[ \t]*$/gim];

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const clean = (s) => (s || "").replace(/\s+/g, " ").trim();
  const lower = (s) => clean(s).toLowerCase();
  const cents = (n) => Math.round(Number(n) * 100);
  const dollars = (c) => (c / 100).toFixed(2);

  const nativeSetter = (el) => {
    const proto = el instanceof HTMLSelectElement ? HTMLSelectElement.prototype
      : el instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    return Object.getOwnPropertyDescriptor(proto, "value").set;
  };

  const isVisible = (el) => !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length);

  /** Find a form control by its label text (prefix match unless exact). */
  function findControl(text, exact = false) {
    const want = lower(text);
    const matches = (s) => (exact ? lower(s) === want : lower(s).startsWith(want));
    for (const label of document.querySelectorAll("label")) {
      if (!matches(label.textContent)) continue;
      const control = label.control
        || (label.htmlFor && document.getElementById(label.htmlFor))
        || label.querySelector("input, select, textarea");
      if (control && isVisible(control)) return control;
    }
    for (const el of document.querySelectorAll("input, select, textarea")) {
      if (matches(el.getAttribute("aria-label")) && isVisible(el)) return el;
    }
    // Older pages: plain text next to the box, no <label>. Take the first box after that text.
    const candidates = [...document.querySelectorAll("td, th, span, div, b, strong, p")]
      .filter((e) => e.children.length === 0 && matches(e.textContent));
    const boxes = [...document.querySelectorAll("input:not([type=hidden]):not([type=button]):not([type=submit]), select, textarea")]
      .filter(isVisible);
    for (const c of candidates) {
      const next = boxes.find((b) => c.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING);
      if (next) return next;
    }
    return null;
  }

  /** Find a button by its text; an exact (case-insensitive) match wins over a partial one. */
  function findButton(text) {
    const want = lower(text);
    const all = [...document.querySelectorAll("button, input[type=button], input[type=submit], [role=button]")]
      .filter(isVisible);
    const textOf = (b) => lower(b.tagName === "INPUT" ? b.value : b.textContent);
    return all.find((b) => textOf(b) === want) || all.find((b) => textOf(b).includes(want)) || null;
  }

  function fire(el, type, init = {}) {
    el.dispatchEvent(new Event(type, { bubbles: true, cancelable: true, ...init }));
  }
  function key(el, type, k) {
    el.dispatchEvent(new KeyboardEvent(type, { key: k, bubbles: true, cancelable: true }));
  }
  /** Leave the box the way a person does (Tab/click away): sites often commit the value on blur. */
  function leave(el) {
    fire(el, "change");
    el.blur();
    fire(el, "focusout");
  }

  /** Type text key by key (keydown/keypress/input/keyup), then change + blur, like a person would. */
  async function typeInto(el, text, delayMs) {
    el.focus();
    const set = nativeSetter(el);
    set.call(el, "");
    fire(el, "input");
    for (const ch of text) {
      key(el, "keydown", ch);
      key(el, "keypress", ch);
      set.call(el, el.value + ch);
      el.dispatchEvent(new InputEvent("input", { bubbles: true, data: ch, inputType: "insertText" }));
      key(el, "keyup", ch);
      if (delayMs) await sleep(delayMs);
    }
    leave(el);
  }

  function setPlain(el, text) {
    el.focus();
    nativeSetter(el).call(el, text);
    fire(el, "input");
    leave(el);
  }

  const numberOf = (s) => Number(String(s).replace(/[$,\s]/g, ""));
  /** Equal as numbers when both are numbers (4 == 4.00), else as trimmed lower-case text. */
  function same(wanted, actual) {
    const a = String(wanted), b = String(actual);
    const na = numberOf(a), nb = numberOf(b);
    if (a.trim() !== "" && b.trim() !== "" && Number.isFinite(na) && Number.isFinite(nb)) return na === nb;
    return lower(a) === lower(b);
  }

  /** The website's running total cost in cents, or null when none is shown yet. */
  function readTotal() {
    const text = document.body.innerText || "";
    for (const pattern of TOTAL_COST_PATTERNS) {
      const found = [...text.matchAll(new RegExp(pattern.source, pattern.flags))];
      if (found.length) return cents(numberOf(found[found.length - 1][1]));
    }
    return null;
  }

  /** Messages the site shows about invalid input (error spans and fields coloured as invalid). */
  function pageErrors() {
    const out = [];
    for (const el of document.querySelectorAll('[id*="ErrorMsg"], [id*="errorMsg"], [role=alert]')) {
      const t = clean(el.textContent);
      if (t && !out.includes(t)) out.push(t);
    }
    for (const el of document.querySelectorAll("input, select")) {
      const bg = (el.style && el.style.backgroundColor) || "";
      if (/pink|rgb\(255, 192, 203\)/i.test(bg)) out.push(`a field is marked invalid (${el.id || el.name || "unnamed"})`);
    }
    return out;
  }

  // --- page probe -------------------------------------------------------------------------
  // A small script placed in the page itself (it needs the page's own JavaScript world) that
  //  * reports whether the site's framework (Oracle ADF) has finished refreshing: data-cf-idle
  //  * reports native alert()/confirm() pop-ups instead of letting them block the page
  // If the site's security policy refuses inline scripts, the probe is silently skipped and
  // the engine falls back to watching the page for changes.
  let probeInstalled = false;
  function installProbe(log) {
    if (probeInstalled) return;
    probeInstalled = true;
    window.addEventListener("message", (e) => {
      if (e.source === window && e.data && e.data.claimFillerDialog) {
        log(`    the page showed a ${e.data.kind} pop-up: ${JSON.stringify(e.data.text)}`);
      }
    });
    try {
      const s = document.createElement("script");
      s.textContent = `(function(){
        if (window.__claimFillerProbe) return; window.__claimFillerProbe = true;
        var root = document.documentElement;
        function idle(){ try { var p = window.AdfPage && window.AdfPage.PAGE;
          return p && p.isSynchronizedWithServer ? String(!!p.isSynchronizedWithServer()) : "unknown"; } catch(e){ return "unknown"; } }
        setInterval(function(){ root.setAttribute("data-cf-idle", idle()); }, 100);
        var say = function(kind, text){ try { window.postMessage({claimFillerDialog:true, kind:kind, text:String(text)}, "*"); } catch(e){} };
        window.alert = function(t){ say("alert", t); };
        window.confirm = function(t){ say("confirm", t); return true; };
      })();`;
      (document.head || document.documentElement).appendChild(s);
      s.remove();
    } catch (e) { /* probe unavailable */ }
  }
  const pageIdle = () => document.documentElement.getAttribute("data-cf-idle"); // "true" | "false" | "unknown" | null

  /** Resolve when the page body has not changed for quietMs (or after maxMs). */
  function quiet(quietMs, maxMs) {
    return new Promise((resolve) => {
      let timer = setTimeout(done, quietMs);
      const stopAt = setTimeout(done, maxMs);
      const obs = new MutationObserver(() => { clearTimeout(timer); timer = setTimeout(done, quietMs); });
      obs.observe(document.body, { subtree: true, childList: true, attributes: true, characterData: true });
      function done() { obs.disconnect(); clearTimeout(timer); clearTimeout(stopAt); resolve(); }
    });
  }

  function create(userOptions = {}) {
    const opt = { ...DEFAULTS, ...userOptions };
    let stopRequested = false;
    const log = userOptions.log || (() => {});
    installProbe(log);

    const boxes = () => ({
      date: findControl(opt.fieldServiceDate),
      service: findControl(opt.fieldService, true),
      hours: findControl(opt.fieldHours, true),
      cost: findControl(opt.fieldCost),
    });

    /** What the four boxes hold right now (the Service box as its visible text). */
    function values() {
      const b = boxes();
      const missing = Object.keys(b).filter((k) => !b[k]);
      if (missing.length) throw new Error(`Could not find the form box(es): ${missing.join(", ")}. Open the Claim details form first.`);
      const sel = b.service;
      return {
        date: b.date.value,
        service: sel.selectedIndex >= 0 ? clean(sel.options[sel.selectedIndex].text) : clean(sel.value),
        hours: b.hours.value,
        cost: b.cost.value,
      };
    }
    function describeBoxes() {
      try {
        const v = values();
        return `form held date=${JSON.stringify(v.date)} service=${JSON.stringify(v.service)} hours=${JSON.stringify(v.hours)} cost=${JSON.stringify(v.cost)}`;
      } catch (e) {
        return `could not read the form (${e.message})`;
      }
    }

    /** Wait until the form stops changing, so a late reset from the previous add can't land mid-typing. */
    async function settle() {
      let last = null;
      for (let i = 0; i < 12; i++) {
        const now = JSON.stringify(values());
        if (now === last) return;
        last = now;
        await sleep(300);
      }
    }

    /**
     * After changing the date or the service the site refreshes parts of the form from the
     * server and may wipe what was typed. Wait out that refresh before typing anything else.
     */
    async function waitForRefresh(why) {
      const started = Date.now();
      await sleep(opt.pprMinPauseMs);
      while (Date.now() - started < opt.pprMaxWaitMs && pageIdle() === "false") await sleep(150);
      const left = Math.max(300, opt.pprMaxWaitMs - (Date.now() - started));
      await quiet(opt.pprQuietMs, left);
      log(`    waited ${Date.now() - started} ms for the page to finish refreshing after the ${why} (page says: ${pageIdle() || "no signal"})`);
    }

    function selectService(name) {
      const sel = boxes().service;
      const want = lower(name);
      const option = [...sel.options].find((o) => lower(o.text) === want);
      if (!option) {
        throw new Error(`Service '${name}' is not in the dropdown (options: ${[...sel.options].map((o) => clean(o.text)).join(", ")}).`);
      }
      sel.focus();
      nativeSetter(sel).call(sel, option.value);
      fire(sel, "input");
      fire(sel, "change");
    }

    /** Put text in a box and check it stuck; re-fill up to fieldTries times. */
    async function setBox(name, text, typed, refreshes) {
      for (let i = 0; i < opt.fieldTries; i++) {
        const el = boxes()[name];
        log(`    ${i ? "re-" : ""}${typed ? "typing" : "setting"} ${name}: ${JSON.stringify(text)}`);
        if (typed) {
          await typeInto(el, text, opt.typingDelayMs);
        } else {
          setPlain(el, text);
          key(el, "keydown", "Escape"); // close a calendar picker if one opened
        }
        if (refreshes) await waitForRefresh(name);
        else await sleep(150);
        const now = values()[name];
        if (same(text, now)) return true;
        log(`    the ${name} box holds ${JSON.stringify(now)} instead of ${JSON.stringify(text)}`);
      }
      return false;
    }

    /**
     * Fill the form for one claim and click Add claim. Returns {status:"clicked"} or
     * {status:"failed", reason, errors}. If the site reloads the page on click, this call is
     * cut off by the reload; the caller then checks the total to see whether the claim landed.
     */
    async function enterClaim(c) {
      try {
        log(`  entering ${c.service_date} ${c.service} ${c.hours}h $${dollars(cents(c.total_cost || 0))}`);
        await settle();
        const wanted = { date: c.service_date, hours: String(c.hours) };
        if (c.total_cost !== null && c.total_cost !== undefined) wanted.cost = dollars(cents(c.total_cost));
        // Date and service are "auto-submit" boxes on the real site: each change makes the page refresh.
        if (!(await setBox("date", wanted.date, false, true))) throw new Error("the service date box would not keep its value");
        log(`    choosing service ${JSON.stringify(c.service)}`);
        selectService(c.service);
        await waitForRefresh("service");
        for (const name of ["hours", "cost"]) {
          if (name in wanted && !(await setBox(name, wanted[name], true, false))) throw new Error(`the ${name} box would not keep its value`);
        }
        // Last look before clicking: a late refresh may have wiped something.
        let ok = false;
        for (let i = 0; i < opt.fieldTries && !ok; i++) {
          const now = values();
          const wrong = Object.keys(wanted).filter((n) => !same(wanted[n], now[n]));
          if (lower(now.service) !== lower(c.service)) wrong.push("service");
          if (!wrong.length) { ok = true; break; }
          log(`    ${wrong.join(", ")} changed while filling (${describeBoxes()}); filling again`);
          for (const n of wrong) {
            if (n === "service") { selectService(c.service); await waitForRefresh("service"); }
            else await setBox(n, wanted[n], n === "hours" || n === "cost", n === "date");
          }
        }
        if (!ok) throw new Error(`form is not right and will not be submitted: ${describeBoxes()}`);
        const errors = pageErrors();
        if (errors.length) throw Object.assign(new Error(`the site shows a problem with the form: ${errors.join(" | ")}`), { errors });
        const add = findButton(opt.buttonAdd);
        if (!add) throw new Error(`could not find the '${opt.buttonAdd}' button`);
        log(`    clicking '${opt.buttonAdd}'`);
        add.click();
        return { status: "clicked" };
      } catch (e) {
        return { status: "failed", reason: e.message, errors: e.errors || pageErrors(), form: describeBoxes() };
      }
    }

    function readState() {
      let form = null;
      let missing = [];
      try { form = values(); } catch (e) { missing = [e.message]; }
      if (!findButton(opt.buttonAdd)) missing.push("add button");
      return { ready: form !== null && missing.length === 0, missing, total: readTotal(), errors: pageErrors(), form, idle: pageIdle() };
    }

    function predetermine() {
      const button = findButton(opt.buttonPredetermine);
      if (!button) return { clicked: false, reason: `could not find the '${opt.buttonPredetermine}' button` };
      log(`  clicking '${opt.buttonPredetermine}'`);
      button.click();
      return { clicked: true };
    }

    async function waitForTotal(expected) {
      const deadline = Date.now() + opt.totalsWaitSeconds * 1000;
      let seen = readTotal();
      while (seen !== expected && Date.now() < deadline) {
        await sleep(250);
        seen = readTotal();
      }
      return seen;
    }

    /**
     * Single-page loop: add each claim and check it landed by comparing the website total with a
     * rolling sum. Only usable when the site does NOT reload on Add claim (the panel runs the
     * loop for the real site). Returns {ok, added, message}.
     */
    async function fill(claims, { predetermine: doPredetermine = true } = {}) {
      stopRequested = false;
      let expected = readTotal() ?? 0;
      log(`  website total cost before adding: $${dollars(expected)}`);
      for (let i = 0; i < claims.length; i++) {
        const c = claims[i];
        const n = i + 1;
        const label = `${c.service_date} ${c.service} ${c.hours}h`;
        const before = expected;
        expected += cents(c.total_cost || 0);
        let added = false;
        for (let attempt = 1; attempt <= opt.maxAttempts && !added; attempt++) {
          if (stopRequested) return { ok: false, added: i, message: `Stopped by you after ${i} of ${claims.length} claims.` };
          const r = await enterClaim(c);
          if (r.status === "clicked") {
            await sleep(300);
            for (let w = Date.now(); Date.now() - w < opt.totalsWaitSeconds * 1000;) {
              const b = findButton(opt.buttonAdd);
              if (b && !b.disabled) break;
              await sleep(200);
            }
          }
          const seen = await waitForTotal(expected);
          if (seen === expected) {
            log(`  added ${n}/${claims.length}: ${label}  (website total now $${dollars(seen)})`);
            added = true;
          } else if (seen !== null && seen !== before) {
            const message = `Stopped at claim ${n} (${label}): the website total is $${dollars(seen)} but $${dollars(expected)} was expected. Check the website before continuing.`;
            log(message);
            return { ok: false, added: i, message };
          } else {
            log(`  claim ${n} (${label}) attempt ${attempt}/${opt.maxAttempts} failed: ${r.reason || "the website total did not change"}; ${describeBoxes()}`);
          }
        }
        if (!added) {
          const message = `Stopped: claim ${n} (${label}) failed ${opt.maxAttempts} times. ${i} of ${claims.length} claims were added.`;
          log(message);
          return { ok: false, added: i, message };
        }
      }
      let message = `All ${claims.length} claims added; total cost matches ($${dollars(expected)}).`;
      if (doPredetermine) {
        const r = predetermine();
        message += r.clicked ? ` Clicked '${opt.buttonPredetermine}'.` : ` But ${r.reason}.`;
        if (!r.clicked) { log("  " + message); return { ok: false, added: claims.length, message }; }
      }
      log("  " + message);
      return { ok: true, added: claims.length, message };
    }

    return { fill, enterClaim, readState, predetermine, stop: () => { stopRequested = true; }, describeBoxes };
  }

  /**
   * One stateless operation for the side panel: "state" | "enter" | "predetermine".
   * Log lines are sent to the panel with runtime messages (the page may reload at any moment).
   */
  async function run(op, payload, options) {
    const log = (text) => {
      try { chrome.runtime.sendMessage({ type: "claim-filler-log", text }); } catch (e) { /* panel closed */ }
    };
    const engine = create({ ...options, log });
    if (op === "state") return engine.readState();
    if (op === "enter") return engine.enterClaim(payload);
    if (op === "predetermine") return engine.predetermine();
    throw new Error(`unknown operation ${op}`);
  }

  globalThis.ClaimFillerEngine = { create, run, readTotal, findControl, findButton, DEFAULTS };
})();
