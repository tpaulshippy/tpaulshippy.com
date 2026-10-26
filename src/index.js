// tpaulshippy.com - Cloudflare Worker fronting static assets.
//
// Static files are served from the ASSETS binding (repo root, minus
// .assetsignore). This script adds canonical-host + legacy-URL redirects
// and security/caching headers on every response.

const CANONICAL_HOST = "tpaulshippy.com";
const BLOG_HOST = "blog.tpaulshippy.com";

// Legacy URLs that must keep working after leaving case-insensitive IIS.
// Keys are matched against the request pathname verbatim.
const LEGACY_ALIASES = new Map([
  // On-disk file is McCain.html; old links sometimes use lowercase.
  ["/mccain.html", "/McCain.html"],
  // tps.htm links ...PiperALibertarian.htm; the file is ...PiperaLibertarian.htm.
  ["/IsJohnPiperALibertarian.htm", "/IsJohnPiperaLibertarian.htm"],
]);

// Blog: retired WordPress routes that collapse to the blog home.
// NOTE: /wp-content/uploads/* is deliberately NOT listed — post images
// live there and must keep serving.
const BLOG_HOME_REDIRECTS = [
  /^\/wp-(admin|login|cron|json|includes)(?:\/|$)/,
  /^\/xmlrpc\.php$/,
  /^\/comments(?:\/|$)/,
  /^\/author(?:\/|$)/,
  /^\/tag(?:\/|$)/,
  /^\/category(?:\/|$)/,
  /^\/page\/\d+\/?$/,
  // Date archives without a post slug (/YYYY/, /YYYY/MM/, /YYYY/MM/DD/).
  /^\/\d{4}\/?$/,
  /^\/\d{4}\/\d{2}\/?$/,
  /^\/\d{4}\/\d{2}\/\d{2}\/?$/,
];

// Fetch an asset with the site-wide security/caching headers applied.
// Extra headers (e.g. CORS) can be passed per response.
async function serveAsset(request, env, extraHeaders) {
  const res = await env.ASSETS.fetch(request);
  const headers = new Headers(res.headers);
  headers.set("X-Content-Type-Options", "nosniff");
  headers.set("Referrer-Policy", "strict-origin-when-cross-origin");
  headers.set("X-Frame-Options", "SAMEORIGIN");
  headers.set(
    "Permissions-Policy",
    "camera=(), microphone=(), geolocation=()",
  );
  if (extraHeaders) {
    for (const [k, v] of Object.entries(extraHeaders)) headers.set(k, v);
  }
  const contentType = res.headers.get("content-type") || "";
  if (contentType.includes("text/html")) {
    headers.set("Cache-Control", "public, max-age=300");
  } else if (res.status === 200) {
    headers.set("Cache-Control", "public, max-age=86400");
  }
  return new Response(res.body, {
    status: res.status,
    statusText: res.statusText,
    headers,
  });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    // www -> apex (keeps one canonical host, preserves path + query).
    if (url.hostname === `www.${CANONICAL_HOST}`) {
      url.hostname = CANONICAL_HOST;
      return Response.redirect(url.toString(), 301);
    }

    // --- Blog host: serve the generated blog/ asset subtree. ---
    if (url.hostname === BLOG_HOST) {
      // Retired WordPress routes collapse to the blog home.
      if (BLOG_HOME_REDIRECTS.some((re) => re.test(url.pathname))) {
        return Response.redirect(`${url.origin}/`, 301);
      }

      // /index.html -> / (canonical home).
      if (url.pathname === "/index.html" || url.pathname === "/index.htm") {
        return Response.redirect(`${url.origin}/`, 301);
      }

      // Map blog URLs into the blog/ asset prefix. /feed/ serves the RSS
      // feed in place so feed readers and the homepage widget keep working.
      let assetPath = `/blog${url.pathname}`;
      if (url.pathname === "/feed" || url.pathname === "/feed/") {
        assetPath = "/blog/feed.xml";
      }
      const allowCors =
        assetPath === "/blog/feed.xml" || assetPath === "/blog/latest.json";

      // Trailing-slash paths resolve to their index.html (permalinks end
      // in / and are stored as .../slug/index.html).
      if (assetPath.endsWith("/")) {
        assetPath += "index.html";
      }
      const assetUrl = new URL(url.origin + assetPath);
      return serveAsset(new Request(assetUrl, request), env, {
        ...(allowCors ? { "Access-Control-Allow-Origin": "*" } : {}),
      });
    }

    // /index.html -> / (canonical home, preserves query).
    if (url.pathname === "/index.html" || url.pathname === "/index.htm") {
      url.pathname = "/";
      return Response.redirect(url.toString(), 301);
    }

    // Legacy case/rename aliases.
    const alias = LEGACY_ALIASES.get(url.pathname);
    if (alias !== undefined) {
      url.pathname = alias;
      return Response.redirect(url.toString(), 301);
    }

    // html_handling = "none" disables directory-index resolution, so map
    // a trailing-slash path to its index.html explicitly (faithful to IIS).
    // The browser URL stays clean (e.g. / serves /index.html with a 200).
    if (url.pathname.endsWith("/")) {
      const indexUrl = new URL(url.origin + url.pathname + "index.html");
      return serveAsset(new Request(indexUrl, request), env);
    }

    return serveAsset(request, env);
  },
};
