// Points the homepage "blog" buttons at the most recent blog post.
//
// Tries the static latest.json first (same content, no backend), then falls
// back to the RSS feed. If both fail (offline, blocked), the fallback hrefs
// to the blog home already in the HTML stay in place.
(function () {
  var BLOG_HOME = "https://blog.tpaulshippy.com/";
  var LATEST_URL = "https://blog.tpaulshippy.com/latest.json";
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
    var fullTitle = title || "Paul's Blog";
    var shortLabel = title || "Paul's Blog";

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

  function fromLatestJson() {
    return fetch(LATEST_URL, { headers: { Accept: "application/json" } }).then(
      function (res) {
        if (!res.ok) throw new Error("latest.json " + res.status);
        return res.json();
      }
    ).then(function (latest) {
      if (!latest || !latest.url) throw new Error("no posts");
      return { title: cleanTitle(latest.title || ""), url: latest.url };
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
    fromLatestJson().then(
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
