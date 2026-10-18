// Points the homepage "blog" buttons at the most recent blog post.
//
// Tries the WordPress REST API first (usually CORS-enabled), then falls back
// to the RSS feed. If both fail (offline, blocked), the fallback hrefs to
// the blog home already in the HTML stay in place.
(function () {
  var BLOG_HOME = "https://blog.tpaulshippy.com/";
  var API_URL =
    "https://blog.tpaulshippy.com/wp-json/wp/v2/posts?per_page=1&_fields=link,title,date";
  var RSS_URL = "https://blog.tpaulshippy.com/feed/";

  function decodeEntities(s) {
    var t = document.createElement("textarea");
    t.innerHTML = s;
    return t.value;
  }

  function stripTags(s) {
    return String(s).replace(/<[^>]*>/g, "");
  }

  function cleanTitle(raw) {
    return decodeEntities(stripTags(raw || "")).trim();
  }

  function applyPost(title, url) {
    if (!url) return;
    var fullTitle = title ? "Latest post: " + title : "Latest post";
    var shortLabel = title ? "Latest: " + title : "Latest post";

    var link = document.getElementById("latest-post-link");
    var label = document.getElementById("latest-post-label");
    if (link) {
      link.href = url;
      link.title = fullTitle;
    }
    if (label) label.textContent = shortLabel;

    var hero = document.getElementById("hero-latest-post");
    var heroLabel = document.getElementById("hero-latest-post-label");
    if (hero) {
      hero.href = url;
      hero.title = fullTitle;
    }
    if (heroLabel) heroLabel.textContent = shortLabel;
  }

  function fromRestApi() {
    return fetch(API_URL, { headers: { Accept: "application/json" } }).then(
      function (res) {
        if (!res.ok) throw new Error("posts API " + res.status);
        return res.json();
      }
    ).then(function (posts) {
      if (!posts || !posts.length) throw new Error("no posts");
      var post = posts[0];
      var title = cleanTitle(post.title && post.title.rendered);
      return { title: title, url: post.link };
    });
  }

  function fromRss() {
    return fetch(RSS_URL, {
      headers: { Accept: "application/rss+xml, application/xml, text/xml" },
    }).then(function (res) {
      if (!res.ok) throw new Error("feed " + res.status);
      return res.text();
    }).then(function (text) {
      var doc = new DOMParser().parseFromString(text, "text/xml");
      var item = doc.querySelector("item");
      if (!item) throw new Error("no items");
      var titleEl = item.querySelector("title");
      var linkEl = item.querySelector("link");
      var url = linkEl ? linkEl.textContent.trim() : "";
      return {
        title: cleanTitle(titleEl ? titleEl.textContent : ""),
        url: url,
      };
    });
  }

  function fetchLatestPost() {
    fromRestApi().then(
      function (viaApi) {
        if (viaApi.url) applyPost(viaApi.title, viaApi.url);
      },
      function () {
        fromRss().then(
          function (viaRss) {
            if (viaRss.url && viaRss.url !== BLOG_HOME)
              applyPost(viaRss.title, viaRss.url);
          },
          function () {
            /* leave fallback blog-home links in place */
          }
        );
      }
    );
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", fetchLatestPost);
  } else {
    fetchLatestPost();
  }
})();
