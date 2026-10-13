// Form-filling engine. Runs inside the claim website's page (injected by the side panel).
// A port of FormFiller in python-app/read_claims.py: it fills one claim at a time and only
// moves on when the website's running total cost equals the rolling sum kept here, retries a
// claim up to 3 times, stops everything after the third failure, and clicks "Predetermine"
// only when every claim in the group is verified.
//
// Plain script (not a module) so it can be injected as-is. Everything lives on
// globalThis.ClaimFillerEngine so tests can drive it without the extension.
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
    totalsWaitSeconds: 5, // how long to wait for the website total to update after Add claim
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
    fire(el, "change");
    el.blur();
    fire(el, "focusout");
  }

  function setPlain(el, text) {
    el.focus();
    nativeSetter(el).call(el, text);
    fire(el, "input");
    fire(el, "change");
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

  function create(userOptions = {}) {
    const opt = { ...DEFAULTS, ...userOptions };
    let stopRequested = false;
    const log = userOptions.log || (() => {});

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
    async function setBox(name, text, typed) {
      for (let i = 0; i < opt.fieldTries; i++) {
        const el = boxes()[name];
        if (typed) {
          await typeInto(el, text, opt.typingDelayMs);
        } else {
          setPlain(el, text);
          key(el, "keydown", "Escape"); // close a calendar picker if one opened
        }
        await sleep(100);
        if (same(text, values()[name])) return true;
      }
      return false;
    }

    async function enterClaim(c) {
      await settle();
      const wanted = { date: c.service_date, hours: String(c.hours) };
      if (c.total_cost !== null && c.total_cost !== undefined) wanted.cost = dollars(cents(c.total_cost));
      if (!(await setBox("date", wanted.date, false))) throw new Error("the service date box would not keep its value");
      selectService(c.service);
      for (const name of ["hours", "cost"]) {
        if (name in wanted && !(await setBox(name, wanted[name], true))) throw new Error(`the ${name} box would not keep its value`);
      }
      // Last look before clicking: a late reset or re-render may have wiped something.
      let ok = false;
      for (let i = 0; i < opt.fieldTries && !ok; i++) {
        const now = values();
        const wrong = Object.keys(wanted).filter((n) => !same(wanted[n], now[n]));
        if (lower(now.service) !== lower(c.service)) wrong.push("service");
        if (!wrong.length) { ok = true; break; }
        log(`    form changed while filling (${describeBoxes()}); filling again`);
        for (const n of wrong) {
          if (n === "service") selectService(c.service);
          else await setBox(n, wanted[n], n === "hours" || n === "cost");
        }
      }
      if (!ok) throw new Error(`form is not right and will not be submitted: ${describeBoxes()}`);
      const add = findButton(opt.buttonAdd);
      if (!add) throw new Error(`could not find the '${opt.buttonAdd}' button`);
      add.click();
      // The site may show "Verifying..." in place of the button while it checks the claim.
      await sleep(300);
      const deadline = Date.now() + opt.totalsWaitSeconds * 1000;
      while (Date.now() < deadline) {
        const b = findButton(opt.buttonAdd);
        if (b && !b.disabled) break;
        await sleep(200);
      }
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
     * Add each claim and check it landed by comparing the website total with a rolling sum.
     * Returns {ok, added, message}. predetermine=false is the "first claim only" test mode.
     */
    async function fill(claims, { predetermine = true } = {}) {
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
          let problem = null;
          try {
            await enterClaim(c);
          } catch (e) {
            problem = `could not use the form (${e.message})`;
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
            log(`  claim ${n} (${label}) attempt ${attempt}/${opt.maxAttempts} failed: ${problem || "the website total did not change"}; ${describeBoxes()}`);
          }
        }
        if (!added) {
          const message = `Stopped: claim ${n} (${label}) failed ${opt.maxAttempts} times. ${i} of ${claims.length} claims were added.`;
          log(message);
          return { ok: false, added: i, message };
        }
      }
      let message = `All ${claims.length} claims added; total cost matches ($${dollars(expected)}).`;
      if (predetermine) {
        const button = findButton(opt.buttonPredetermine);
        if (!button) {
          message += ` But the '${opt.buttonPredetermine}' button could not be found.`;
          log("  " + message);
          return { ok: false, added: claims.length, message };
        }
        button.click();
        message += ` Clicked '${opt.buttonPredetermine}'.`;
      }
      log("  " + message);
      return { ok: true, added: claims.length, message };
    }

    /** Is the Claim details form on this page? Lists what is missing. */
    function check() {
      const b = boxes();
      const missing = Object.keys(b).filter((k) => !b[k]);
      if (!findButton(opt.buttonAdd)) missing.push("add button");
      return { ready: missing.length === 0, missing, total: readTotal() };
    }

    return { fill, check, stop: () => { stopRequested = true; }, describeBoxes };
  }

  globalThis.ClaimFillerEngine = { create, readTotal, findControl, findButton, DEFAULTS };

  // In the extension: talk to the side panel over a port.
  if (typeof chrome !== "undefined" && chrome.runtime && chrome.runtime.onConnect) {
    chrome.runtime.onConnect.addListener((port) => {
      if (port.name !== "claim-filler") return;
      let engine = null;
      port.onMessage.addListener(async (msg) => {
        try {
          if (msg.type === "check") {
            port.postMessage({ type: "checked", ...create(msg.options).check() });
          } else if (msg.type === "fill") {
            engine = create({ ...msg.options, log: (text) => port.postMessage({ type: "log", text }) });
            const outcome = await engine.fill(msg.claims, { predetermine: msg.predetermine });
            port.postMessage({ type: "done", outcome });
          } else if (msg.type === "stop" && engine) {
            engine.stop();
          }
        } catch (e) {
          port.postMessage({ type: "done", outcome: { ok: false, added: 0, message: `Stopped: ${e.message}` } });
        }
      });
    });
  }
})();
