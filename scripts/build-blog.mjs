// Builds the static blog from blog-src/posts/*.md into blog/.
//
// Sources: Markdown with frontmatter (title, date, slug).
// Outputs (all under blog/, git-ignored, built in CI and locally):
//   blog/index.html                      blog home (newest first)
//   blog/<yyyy>/<mm>/<dd>/<slug>/       one dir per post (matches WP permalinks)
//   blog/feed.xml                        RSS 2.0 (full content)
//   blog/latest.json                     { title, url, date } for latest-post.js
//   blog/sitemap.xml
//   blog/wp-content/uploads/...         copied from blog-src/uploads
// Media referenced as /wp-content/uploads/... resolves on the blog host
// because the worker serves blog URLs from the blog/ asset prefix.

import { readdirSync, readFileSync, writeFileSync, mkdirSync, cpSync, rmSync, existsSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..");
const SRC = join(ROOT, "blog-src", "posts");
const UPLOADS_SRC = join(ROOT, "blog-src", "uploads");
const OUT = join(ROOT, "blog");
const BLOG_HOME = "https://blog.tpaulshippy.com";

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function escapeAttr(s) {
  return escapeHtml(s).replace(/\n/g, " ");
}

// --- Minimal Markdown renderer (subset used by the posts) -----------------

function renderInline(s) {
  // stash code spans
  const codes = [];
  s = s.replace(/`([^`]+?)`/g, (_, c) => {
    codes.push(`<code>${escapeHtml(c)}</code>`);
    return `\ue000${codes.length - 1}\ue001`;
  });
  // images
  s = s.replace(/!\[([^\]]*)\]\(([^)\s]+)\)/g, (_, alt, src) => {
    const a = escapeAttr(alt);
    if (a) return `<figure><img src="${src}" alt="${a}" loading="lazy" /><figcaption>${a}</figcaption></figure>`;
    return `<img src="${src}" alt="" loading="lazy" />`;
  });
  // links
  s = s.replace(/\[([^\]]+)\]\(([^)\s]+)\)/g, '<a href="$2">$1</a>');
  // bold, italic (after links so link text markup survives)
  s = s.replace(/\*\*([^*]+?)\*\*/g, "<strong>$1</strong>");
  s = s.replace(/(^|[^*\w])\*([^*\n]+?)\*/g, "$1<em>$2</em>");
  // restore code
  s = s.replace(/(\d+)/g, (_, i) => codes[Number(i)]);
  return s;
}

function renderMarkdown(md) {
  const lines = md.split("\n");
  let html = "";
  let i = 0;
  let inCode = false;
  let codeBuf = [];
  let listTag = null; // 'ul' | 'ol'
  let quoteBuf = null;

  function closeList() {
    if (listTag) {
      html += `</${listTag}>\n`;
      listTag = null;
    }
  }
  function closeQuote() {
    if (quoteBuf) {
      html += `<blockquote>\n${quoteBuf.join("\n")}\n</blockquote>\n`;
      quoteBuf = null;
    }
  }

  const flushPara = (buf) => {
    if (buf.length) html += `<p>${renderInline(buf.join(" "))}</p>\n`;
  };
  let para = [];

  for (const raw of lines) {
    const line = raw.replace(/\s+$/, "");
    if (inCode) {
      if (/^```/.test(line)) {
        html += `<pre><code>${escapeHtml(codeBuf.join("\n"))}</code></pre>\n`;
        codeBuf = [];
        inCode = false;
      } else {
        codeBuf.push(raw.replace(/\s+$/, ""));
      }
      continue;
    }
    if (/^```/.test(line)) {
      flushPara(para); para = [];
      closeList(); closeQuote();
      inCode = true;
      continue;
    }
    if (/^\s*$/.test(line)) {
      flushPara(para); para = [];
      closeList(); closeQuote();
      continue;
    }
    const h = line.match(/^(#{1,6})\s+(.*)$/);
    if (h) {
      flushPara(para); para = [];
      closeList(); closeQuote();
      html += `<h${h[1].length}>${renderInline(h[2])}</h${h[1].length}>\n`;
      continue;
    }
    if (/^---\s*$/.test(line)) {
      flushPara(para); para = [];
      closeList(); closeQuote();
      html += "<hr />\n";
      continue;
    }
    const bq = line.match(/^>\s?(.*)$/);
    if (bq) {
      flushPara(para); para = [];
      closeList();
      if (!quoteBuf) quoteBuf = [];
      quoteBuf.push(`<p>${renderInline(bq[1])}</p>`);
      continue;
    }
    const ul = line.match(/^[-*]\s+(.*)$/);
    const ol = line.match(/^\d+\.\s+(.*)$/);
    if (ul || ol) {
      flushPara(para); para = [];
      closeQuote();
      const tag = ul ? "ul" : "ol";
      if (listTag !== tag) {
        closeList();
        html += `<${tag}>\n`;
        listTag = tag;
      }
      html += `<li>${renderInline((ul || ol)[1])}</li>\n`;
      continue;
    }
    closeList(); closeQuote();
    para.push(line.trim());
  }
  flushPara(para);
  closeList(); closeQuote();
  if (inCode) html += `<pre><code>${escapeHtml(codeBuf.join("\n"))}</code></pre>\n`;
  return html;
}

// --- Page templates (match the main site aesthetic) -----------------------

function head(title, description, canonical) {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>${escapeHtml(title)}</title>
  <meta name="description" content="${escapeAttr(description)}" />
  <link rel="canonical" href="${canonical}" />
  <link rel="shortcut icon" type="image/x-icon" href="https://tpaulshippy.com/favicon.ico" />
  <link rel="alternate" type="application/rss+xml" title="Paul's Blog" href="/feed.xml" />
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet" />
  <link href="https://tpaulshippy.com/style.css?v=4" type="text/css" rel="stylesheet" />
  <style>
    .post-content img { max-width: 100%; height: auto; border-radius: 8px; }
    .post-content figure { margin: 1.5em 0; }
    .post-content figcaption { font-size: 0.85em; opacity: 0.75; margin-top: 0.4em; }
    .post-content pre { overflow-x: auto; padding: 1em; border-radius: 8px; }
    .post-meta { opacity: 0.75; font-size: 0.9em; margin-bottom: 1.5em; }
  </style>
</head>
<body>
  <a href="#main" class="skip-link">Skip to content</a>
  <header class="site-header">
    <div class="header-inner">
      <a href="https://tpaulshippy.com/" class="brand">
        <span class="avatar avatar-sm"><span class="avatar-fallback">PS</span></span>
        <span class="brand-text">
          <span class="brand-name">Paul's Blog</span>
          <span class="brand-sub">TPaulShippy.com</span>
        </span>
      </a>
      <nav class="main-nav" aria-label="Primary">
        <a class="btn btn-ghost btn-sm" href="https://tpaulshippy.com/">Home</a>
        <a class="btn btn-ghost btn-sm" href="/">Blog</a>
        <a class="btn btn-ghost btn-sm" href="/feed.xml">RSS</a>
      </nav>
    </div>
  </header>
  <main id="main" class="container">`;
}

const FOOT = `  </main>
  <footer class="site-footer">
    <p>&copy; <span id="year"></span> Paul Shippy &middot; <a class="link" href="https://tpaulshippy.com/">TPaulShippy.com</a></p>
  </footer>
  <script>document.getElementById('year').textContent = new Date().getFullYear();</script>
</body>
</html>
`;

function postPage(post) {
  const tags = post.tags.length
    ? `<p class="post-meta">Tagged: ${post.tags.map(escapeHtml).join(", ")}</p>`
    : "";
  return `${head(`${post.title} — Paul's Blog`, post.excerpt, post.url)}
    <section class="hero-section">
      <div class="card">
        <div class="card-content">
          <p class="eyebrow">Paul's Blog</p>
          <h1 class="hero-title">${escapeHtml(post.title)}</h1>
          <p class="post-meta"><time datetime="${post.date}">${post.displayDate}</time> &middot; Learning</p>
          <div class="post-content">${post.html}</div>
          ${tags}
          <div class="hero-cta">
            <a class="btn btn-outline" href="/">All posts</a>
            <a class="btn btn-ghost" href="https://tpaulshippy.com/">tpaulshippy.com</a>
          </div>
        </div>
      </div>
    </section>
${FOOT}`;
}

function indexPage(posts) {
  const items = posts
    .map(
      (p) => `            <article class="repo-card">
              <div class="repo-head">
                <a class="repo-name" href="${p.path}">${escapeHtml(p.title)}</a>
              </div>
              <p class="repo-desc">${escapeHtml(p.excerpt)}</p>
              <div class="repo-meta"><span><time datetime="${p.date}">${p.displayDate}</time></span></div>
            </article>`
    )
    .join("\n");
  return `${head("Paul's Blog — Christian software developer notes", "Notes on software, homesteading, and learning by Paul Shippy.", `${BLOG_HOME}/`)}
    <section class="hero-section">
      <div class="card">
        <div class="card-content">
          <p class="eyebrow">Paul's Blog</p>
          <h1 class="hero-title">Notes &amp; Essays</h1>
          <p class="hero-sub">Short posts on software, learning, and life. Newest first.</p>
        </div>
      </div>
    </section>
    <section class="cards-grid">
      <div class="card">
        <div class="card-content">
          <div class="stack">
${items}
          </div>
        </div>
      </div>
    </section>
${FOOT}`;
}

// --- Build ---------------------------------------------------------------

function parsePost(text, file) {
  const m = text.match(/^---\n([\s\S]*?)\n---\n([\s\S]*)$/);
  if (!m) throw new Error(`bad frontmatter: ${file}`);
  const meta = {};
  for (const line of m[1].split("\n")) {
    const kv = line.match(/^(\w+):\s*(.*)$/);
    if (kv) meta[kv[1]] = kv[2].replace(/^"|"$/g, "");
  }
  // tags: [a, b] -> array
  if (typeof meta.tags === "string") {
    meta.tags = meta.tags.replace(/^\[|\]$/g, "").split(",").map((t) => t.trim()).filter(Boolean);
  } else {
    meta.tags = [];
  }
  return { meta, body: m[2].trim() + "\n" };
}

function excerptOf(html) {
  const text = html.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
  return text.length > 160 ? text.slice(0, 157).trimEnd() + "…" : text;
}

function main() {
  rmSync(OUT, { recursive: true, force: true });
  mkdirSync(OUT, { recursive: true });

  const files = readdirSync(SRC).filter((f) => f.endsWith(".md")).sort();
  if (!files.length) throw new Error("no posts found in blog-src/posts");
  const posts = files.map((file) => {
    const { meta, body } = parsePost(readFileSync(join(SRC, file), "utf8"), file);
    if (!meta.title || !meta.date || !meta.slug) throw new Error(`missing frontmatter in ${file}`);
    const [y, mo, d] = meta.date.split("-");
    const path = `/${y}/${mo}/${d}/${meta.slug}/`;
    const html = renderMarkdown(body);
    const displayDate = new Date(`${meta.date}T12:00:00`).toLocaleDateString("en-US", {
      year: "numeric", month: "long", day: "numeric",
    });
    return { ...meta, path, html, url: BLOG_HOME + path, excerpt: excerptOf(html), displayDate };
  });
  posts.sort((a, b) => (a.date < b.date ? 1 : -1));

  // post pages (permalink dirs preserved)
  for (const p of posts) {
    const dir = join(OUT, p.path);
    mkdirSync(dir, { recursive: true });
    writeFileSync(join(dir, "index.html"), postPage(p));
  }
  // home
  writeFileSync(join(OUT, "index.html"), indexPage(posts));

  // RSS feed (full content)
  const rssItems = posts
    .map((p) => `    <item>
      <title>${escapeHtml(p.title)}</title>
      <link>${p.url}</link>
      <guid isPermaLink="true">${p.url}</guid>
      <pubDate>${new Date(`${p.date}T12:00:00Z`).toUTCString()}</pubDate>
      <description>${escapeHtml(p.excerpt)}</description>
      <content:encoded><![CDATA[${p.html.replaceAll("]]>", "]]&gt;")}]]></content:encoded>
    </item>`)
    .join("\n");
  writeFileSync(
    join(OUT, "feed.xml"),
    `<?xml version="1.0" encoding="UTF-8"?>\n<rss version="2.0" xmlns:content="http://purl.org/rss/1.0/modules/content/">\n  <channel>\n    <title>Paul's Blog</title>\n    <link>${BLOG_HOME}/</link>\n    <description>Notes on software, homesteading, and learning by Paul Shippy.</description>\n    <language>en-us</language>\n${rssItems}\n  </channel>\n</rss>\n`
  );

  // latest.json for the homepage widget
  const [latest] = posts;
  writeFileSync(
    join(OUT, "latest.json"),
    JSON.stringify({ title: latest.title, url: latest.url, date: latest.date })
  );

  // sitemap
  const smUrls = [`  <url><loc>${BLOG_HOME}/</loc></url>`].concat(
    posts.map((p) => `  <url><loc>${p.url}</loc><lastmod>${p.date}</lastmod></url>`)
  );
  writeFileSync(
    join(OUT, "sitemap.xml"),
    `<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n${smUrls.join("\n")}\n</urlset>\n`
  );

  // uploads (paths preserved exactly)
  if (existsSync(UPLOADS_SRC)) {
    cpSync(UPLOADS_SRC, join(OUT, "wp-content", "uploads"), { recursive: true });
  }

  console.log(`blog build: ${posts.length} posts -> ${OUT}`);
}

main();
