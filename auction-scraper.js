/**
 * CSULB Auction Console — listings scraper
 * =========================================
 * Tiny Cloudflare Worker that fetches the public Public Surplus and GovDeals
 * search pages server-side and returns the raw HTML as JSON. The console's
 * client-side parsers (parsePublicSurplusList / parseGovDealsList in
 * features.js) do the actual extraction — the Worker just gets around the
 * browser's CORS / anti-bot restrictions by fetching from Cloudflare's
 * network instead of the user's browser.
 *
 * Deploy:
 *   npm install -g wrangler
 *   wrangler login
 *   wrangler deploy            # from this directory
 *
 * The deploy will print a URL like
 *   https://csulb-auction-scraper.<your-subdomain>.workers.dev
 *
 * Paste that URL (no trailing slash) into the Pull Listings modal's
 * "Scraper Worker URL" field in the console, then click Fetch from web.
 *
 * No KV, no secrets — the Worker has no state and only reads two public
 * web pages. Cloudflare's edge cache (60s) keeps it friendly to the origin.
 */

const PS_URL =
  "https://www.publicsurplus.com/sms/list/current?orgid=1228";
const GD_URL =
  "https://www.govdeals.com/en/search?accountId=24860&companyName=California%20State%20University%20-%20Long%20Beach,%20CA";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "*",
};

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "Content-Type": "application/json", ...CORS },
  });
}

async function fetchHtml(url) {
  const res = await fetch(url, {
    headers: {
      // Look like a normal browser so we don't get a soft-block
      "User-Agent":
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36",
      Accept: "text/html,application/xhtml+xml",
      "Accept-Language": "en-US,en;q=0.9",
    },
    cf: { cacheTtl: 60, cacheEverything: true },
  });
  if (!res.ok) throw new Error("HTTP " + res.status);
  return res.text();
}

export default {
  async fetch(request) {
    const url = new URL(request.url);

    if (request.method === "OPTIONS") {
      return new Response(null, { headers: CORS });
    }

    if (url.pathname === "/" || url.pathname === "") {
      return new Response(
        "CSULB Auction Console scraper. GET /listings?platform=ps|gd|both",
        { headers: { "Content-Type": "text/plain" } }
      );
    }

    if (url.pathname !== "/listings") {
      return json({ error: "Not found. Use /listings?platform=ps|gd|both" }, 404);
    }

    const platform = (url.searchParams.get("platform") || "both").toLowerCase();
    const out = { PS: "", GD: "", errors: [] };

    if (platform === "ps" || platform === "both") {
      try {
        out.PS = await fetchHtml(PS_URL);
      } catch (e) {
        out.errors.push("Public Surplus: " + e.message);
      }
    }
    if (platform === "gd" || platform === "both") {
      try {
        out.GD = await fetchHtml(GD_URL);
      } catch (e) {
        out.errors.push("GovDeals: " + e.message);
      }
    }

    return json(out);
  },
};