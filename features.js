/* ===================================================================
   features.js — CSULB Auction Console v6 add-ons
   - parseListingHtml()  : read a pasted Public Surplus / GovDeals
                           listing description block into structured fields
   - convertListing()    : rewrite a listing block from one platform's
                           HTML conventions into the other's
   - pullListings()      : best-effort fetch of the two public search
                           pages through a CORS proxy (paste is the
                           reliable fallback; see note in the modal)
   - buildRecap()        : generate a weekly recap in David's email style
   These functions are attached to window so app.js / index.html can call
   them; everything else in the app stays as-is.
   =================================================================== */
(function () {
  "use strict";

  /* ---------- shared helpers ---------- */
  function txt(el) { return (el ? el.textContent : "").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim(); }
  function firstMatch(s, re) { const m = s.match(re); return m ? m[1].trim() : ""; }

  /* ===================================================================
     1. PARSE a pasted listing description (PS or GD)
     Returns { title, condition, survey, tag, platform, bullets[], notes }
     Works on either platform's HTML — they differ in tags but the
     meaningful content (heading, bullets, "Survey# / Tag#",
     "Condition: FAIR") is the same.
     =================================================================== */
  function parseListingHtml(html) {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const flat = txt(doc.body);

    // Heading: PS wraps it in <span underline><strong>…, GD in <b><u>…
    // Strategy: the first <strong>/<b> that sits above a <ul> and whose
    // text looks like a product name (not the red payment boilerplate).
    let title = "";
    const heads = [...doc.querySelectorAll("strong, b, u")];
    for (const h of heads) {
      const t = txt(h);
      if (!t) continue;
      if (/payment\s+due|removal\s+required|important|disclaimer|standard|sold as-is|as-is|^as is/i.test(t)) continue;
      if (/^(condition|fair|good|poor|new|excellent)\b/i.test(t)) continue;
      if (t.length >= 4 && t.length <= 120) { title = t; break; }
    }

    // Condition: "Condition: FAIR" (GD) or "...condition is listed as FAIR" (PS)
    let condition = firstMatch(flat, /condition[:\s]+is\s+listed\s+as\s+([A-Za-z]+)/i)
      || firstMatch(flat, /condition[:\s]+([A-Za-z]+)/i);
    condition = condition ? condition.toUpperCase() : "";
    if (/(GOOD|FAIR|POOR|NEW|EXCELLENT|SALVAGE)/.test(condition) === false) {
      // last resort: scan for a lone condition word near "as is"
      condition = (firstMatch(flat, /\b(FAIR|GOOD|POOR|NEW|EXCELLENT|SALVAGE)\b/) || "").toUpperCase();
    }

    // Internal fields
    const survey = firstMatch(flat, /survey\s*#?\s*([0-9]{2}-?[0-9]{3,4})/i);
    const tag = firstMatch(flat, /tag\s*#?\s*([0-9]{4,6})/i);

    // Platform guess from markup conventions
    const lower = html.toLowerCase();
    let platform = "";
    if (/text-align:\s*center;\s*padding-left/.test(lower) || /<ul>\s*<ul>/.test(lower)) platform = "PS";
    if (/<b><u>|<li><b>/.test(lower)) platform = platform || "GD";

    // Bullets (item details), excluding the red payment lines
    const bullets = [...doc.querySelectorAll("li")]
      .map(txt)
      .filter(t => t && !/payment\s+due|removal\s+required/i.test(t));

    return { title, condition, survey, tag, platform, bullets, notes: flat };
  }

  /* ===================================================================
     2. CONVERT a listing block between platforms
     PS markup  ->  GD markup, or GD -> PS.
     We rebuild from the parsed structure so the output is clean and
     consistent rather than a fragile tag-by-tag rewrite.
     =================================================================== */
  function convertListing(html, target /* "GD" | "PS" */) {
    const p = parseListingHtml(html);
    // Re-extract nested bullets to preserve the two-level structure
    const doc = new DOMParser().parseFromString(html, "text/html");
    const topLis = [...doc.querySelectorAll("ul > li")].filter(li => {
      const t = txt(li);
      return t && !/payment\s+due|removal\s+required/i.test(t);
    });

    function liTree(ul) {
      // returns array of {text, children:[...]}
      // Handles BOTH conventions:
      //   GD: nested <ul> lives INSIDE the <li>
      //   PS: nested <ul> is a SIBLING that follows the <li>
      const out = [];
      const kids = [...ul.children];
      for (let i = 0; i < kids.length; i++) {
        const node = kids[i];
        if (node.tagName === "LI") {
          const sub = node.querySelector(":scope > ul");
          let own = txt(node);
          if (sub) own = own.replace(txt(sub), "").trim();
          let children = sub ? liTree(sub) : [];
          // PS pattern: a <ul> sibling immediately after this <li>
          if (!sub && kids[i + 1] && kids[i + 1].tagName === "UL") {
            children = liTree(kids[i + 1]);
            i++; // consume the sibling ul
          }
          out.push({ text: own, children });
        } else if (node.tagName === "UL" && !out.length) {
          // leading stray ul with no preceding li
          out.push(...liTree(node));
        }
      }
      return out;
    }
    const firstUl = doc.querySelector("ul");
    const tree = firstUl ? liTree(firstUl) : p.bullets.map(b => ({ text: b, children: [] }));

    const esc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    const redBoiler = isPS =>
      isPS
        ? `<p style="text-align: center; padding-left: 30px;"><span style="color: #ff0000;"><strong><em>IMPORTANT: Please review our Standard Disclaimer and Terms and Conditions. ====&gt;</em></strong></span></p>\n<p style="text-align: center; padding-left: 30px;"><span style="color: #ff0000;"><strong>PAYMENT&nbsp;DUE WITHIN 5 BUSINESS DAYS OF AUCTION CLOSE.<br /><br />REMOVAL REQUIRED WITHIN 10 BUSINESS DAYS OF PAYMENT POSTED.<br /></strong></span></p>`
        : `<p style="text-align: center; color: #ff0000;"><strong><em>IMPORTANT: Please review our Standard Disclaimer and Terms and Conditions. </em></strong></p>\n<p style="text-align: center; color: #ff0000;"><strong>PAYMENT DUE WITHIN 5 BUSINESS DAYS OF AUCTION CLOSE.<br>REMOVAL REQUIRED WITHIN 10 BUSINESS DAYS OF PAYMENT POSTED.</strong></p>`;

    function renderListPS(nodes) {
      let h = "<ul>\n";
      for (const n of nodes) {
        h += `<li>${esc(n.text)}</li>\n`;
        if (n.children.length) h += renderListPS(n.children);
      }
      return h + "</ul>\n";
    }
    function renderListGD(nodes) {
      // GD sample omits closing </li> and uses <b> on top-level labels
      let h = "<ul>\n";
      for (const n of nodes) {
        const top = nodes === tree;
        h += `  <li>${top ? "<b>" + esc(n.text) + "</b>" : esc(n.text)}\n`;
        if (n.children.length) h += renderListGD(n.children);
      }
      return h + "</ul>\n";
    }

    const title = p.title || "Item";
    const cond = p.condition || "FAIR";

    if (target === "PS") {
      return [
        redBoiler(true),
        `<br />`,
        `<p><span style="text-decoration: underline;"><strong>${esc(title)}</strong></span></p>`,
        renderListPS(tree),
        `<br /><br /><br /> While the condition is listed as<strong>&nbsp;${esc(cond)}</strong>, the item is being sold "<strong>As is, Where-is</strong>". The State reserves the right to cancel the bid at any time and makes no warranty as to the expressed condition of the equipment. <br /><br /><br />`,
        `<div><span style="text-decoration: underline;">~For Internal Use~</span></div>`,
        `&nbsp;Survey#&nbsp;${esc(p.survey || "")}<br /> &nbsp;Tag#&nbsp;${esc(p.tag || "")}`
      ].join("\n");
    }
    // target === "GD"
    return [
      redBoiler(false),
      `<b><u>${esc(title)}</u></b>`,
      renderListGD(tree),
      `<p>Condition: <strong>${esc(cond)}</strong></p>`,
      `<p>This item is sold "As is, Where-is". The State reserves the right to cancel the bid at any time and makes no warranty as to the expressed condition of the equipment.</p>`,
      `<u>~For Internal Use~</u>`,
      `Survey# ${esc(p.survey || "")}`,
      `Tag# ${esc(p.tag || "")}`
    ].join("\n");
  }

  /* ===================================================================
     3. PULL public listings through a CORS proxy.
     A static page opened from disk cannot fetch GovDeals / Public
     Surplus directly (the sites block non-browser requests and the
     browser blocks cross-origin reads). A public CORS proxy is the
     only way to do it without a backend; it is best-effort and may be
     rate-limited or down, which is exactly why paste-the-HTML is the
     primary, always-works path. We surface failures loudly rather than
     returning nothing.
     =================================================================== */
  const PROXY = "https://corsproxy.io/?url=";
  const PS_URL = "https://www.publicsurplus.com/sms/list/current?orgid=1228";
  const GD_URL = "https://www.govdeals.com/en/search?accountId=24860&companyName=California%20State%20University%20-%20Long%20Beach,%20CA";

  async function fetchVia(url) {
    const res = await fetch(PROXY + encodeURIComponent(url), { headers: { "Accept": "text/html" } });
    if (!res.ok) throw new Error("HTTP " + res.status);
    return res.text();
  }

  function parsePublicSurplusList(html) {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const out = [];
    // PS current-listing table rows link to /sms/<org>/auction/view?auc=<id>
    doc.querySelectorAll('a[href*="auction/view"]').forEach(a => {
      const title = txt(a);
      if (!title) return;
      const row = a.closest("tr") || a.parentElement;
      const rowText = txt(row);
      const bid = firstMatch(rowText, /\$([\d,]+(?:\.\d{2})?)/);
      out.push({ platform: "PS", title, url: new URL(a.getAttribute("href"), PS_URL).href, currentBid: bid });
    });
    return dedupe(out);
  }

  function parseGovDealsList(html) {
    const out = [];
    // GovDeals is a Next.js app; the listing payload rides in __NEXT_DATA__
    const doc = new DOMParser().parseFromString(html, "text/html");
    const nd = doc.querySelector("#__NEXT_DATA__");
    if (nd) {
      try {
        const data = JSON.parse(nd.textContent);
        const stack = [data];
        while (stack.length) {
          const cur = stack.pop();
          if (cur && typeof cur === "object") {
            if (Array.isArray(cur)) cur.forEach(v => stack.push(v));
            else {
              if (cur.title && (cur.currentPrice != null || cur.currentBid != null || cur.id)) {
                out.push({
                  platform: "GD",
                  title: String(cur.title),
                  url: cur.id ? GD_URL.split("/en/")[0] + "/en/asset/" + cur.id : "",
                  currentBid: String(cur.currentPrice != null ? cur.currentPrice : (cur.currentBid != null ? cur.currentBid : ""))
                });
              }
              Object.values(cur).forEach(v => { if (v && typeof v === "object") stack.push(v); });
            }
          }
        }
      } catch (e) { /* fall through to DOM */ }
    }
    if (!out.length) {
      doc.querySelectorAll('a[href*="/asset/"]').forEach(a => {
        const title = txt(a);
        if (title) out.push({ platform: "GD", title, url: new URL(a.getAttribute("href"), GD_URL).href, currentBid: "" });
      });
    }
    return dedupe(out);
  }

  function dedupe(arr) {
    const seen = new Set(), out = [];
    for (const x of arr) { const k = x.platform + "|" + x.title.toLowerCase(); if (!seen.has(k)) { seen.add(k); out.push(x); } }
    return out;
  }

  async function pullListings() {
    const results = { PS: [], GD: [], errors: [] };
    try { results.PS = parsePublicSurplusList(await fetchVia(PS_URL)); }
    catch (e) { results.errors.push("Public Surplus: " + e.message); }
    try { results.GD = parseGovDealsList(await fetchVia(GD_URL)); }
    catch (e) { results.errors.push("GovDeals: " + e.message); }
    return results;
  }

  /* ===================================================================
     4. RECAP generator — mirrors David's weekly email.
     Sections: Sold / Collected / Prepping (Relist, New, Released) and
     groups items by Department, with "Item – Asking Price $X".
     Pulls live items[] from the app (passed in).
     =================================================================== */
  function money(v) {
    const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ""));
    return isNaN(n) ? "" : "$" + n.toLocaleString("en-US");
  }
  function askingPrice(it) {
    // prefer an explicit list price, else current bid, else est value
    const cands = [it.platforms && it.platforms.GD && it.platforms.GD.price,
                   it.platforms && it.platforms.PS && it.platforms.PS.price,
                   it.currentBid, it.amount];
    for (const c of cands) { const m = money(c); if (m) return m; }
    return "";
  }
  function line(it) {
    const ap = askingPrice(it);
    return "*  " + it.description + (ap ? " – Asking Price " + ap : "");
  }
  function groupByDept(list) {
    const m = new Map();
    for (const it of list) {
      const d = (it.dept || "Other").trim() || "Other";
      if (!m.has(d)) m.set(d, []);
      m.get(d).push(it);
    }
    return [...m.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  }
  function section(list) {
    if (!list.length) return "";
    let s = "";
    for (const [dept, arr] of groupByDept(list)) {
      s += "o   " + dept + "\n";
      for (const it of arr) s += line(it) + "\n";
    }
    return s;
  }

  function buildRecap(items, opts) {
    opts = opts || {};
    const today = new Date();
    const wk = opts.weekOf || (today.getMonth() + 1) + "/" + today.getDate();
    const to = opts.to || "Sam";
    const from = opts.from || "David Pedersen";

    const sold = items.filter(i => ["SOLD", "PAID"].includes(i.status));
    const collected = items.filter(i => i.status === "PICKED_UP");
    const relist = items.filter(i => i.status === "RELISTED" || i.status === "UNSOLD");
    const isNew = items.filter(i => i.status === "PREP");
    const released = items.filter(i => i.status === "LIVE");

    let out = "";
    out += "Hi " + to + ",\n\n\nHere is the recap – week of " + wk + ".\n\n";
    out += "Auctions\n\n\n";

    out += "Sold\n";
    out += sold.length ? section(sold) : "*\tNo new sales\n";
    out += "\n\nCollected\n";
    out += collected.length ? section(collected) : "*\tNo auctions collected\n";

    out += "\n\nPrepping\n\n\n";
    out += "Relist\n\n";
    out += relist.length ? section(relist) : "(none)\n";
    out += "\nNew\n\n";
    out += isNew.length ? section(isNew) : "(none)\n";
    out += "\nReleased\n\n";
    out += released.length ? section(released) : "(none)\n";

    out += "\n\nOther\n\n";
    out += "*\tDay to day - Property Taggings, Survey Processing, L&F Inquiries, Surplus Furniture Inquiries\n\n\n";
    out += "Thank you,\n\n" + from + "\nProperty Assistant\nCalifornia State University, Long Beach\n";
    return out;
  }

  /* expose */
  window.AuctionFeatures = {
    parseListingHtml, convertListing, pullListings,
    parsePublicSurplusList, parseGovDealsList, buildRecap
  };
})();
