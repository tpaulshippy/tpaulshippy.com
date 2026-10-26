---
title: "Bitwarden subdomains"
date: 2020-07-29
slug: bitwarden-subdomains
---

#### Context

I use Bitwarden as my password manager. It is open source, battle tested, and secure. But I realized recently that it doesn’t recognize a subdomain as a separate website to which I can login. One of the applications I use daily has many instances that are hosted on subdomains and each instance has its own username and password. When I need to access an instance, I have to navigate to the URL and then manually select the right option from a long list in my Bitwarden plugin for Brave to autofill my username and password. This got me thinking – I wonder if Bitwarden has an option to handle this situation.

#### Search terms

bitwarden subdomain

#### Helpful link

[https://bitwarden.com/help/article/uri-match-detection/](https://bitwarden.com/help/article/uri-match-detection/)

#### Outcome

Turns out the Bitwarden team has solved this with a feature. By changing my default URI match detection option to “Host” I get the behavior I was hoping for. The plugin matches the right instance by subdomain and provides me a quick way to fill in my username and password.
