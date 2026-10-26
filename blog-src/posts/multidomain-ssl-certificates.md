---
title: "Multidomain SSL Certificates"
date: 2020-06-15
slug: multidomain-ssl-certificates
---

#### Context

As I relaunch my website and blog, I want to make it secure. But I host two different domains on this same site (with a host-based redirect), so I wanted to make sure I could secure both of them. My host’s support team told me I would need to use their “Certificate Signing Request” (CSR) to install an SSL cert but this requires a single common name for the domain name. My question was – how do I generate a CSR for multiple domains?

#### Search terms

csr for multidomain certificate

#### Helpful link

[https://www.namecheap.com/support/knowledgebase/article.aspx/9840/67/how-do-i-activate-a-multidomain-ssl-certificate](https://www.namecheap.com/support/knowledgebase/article.aspx/9840/67/how-do-i-activate-a-multidomain-ssl-certificate)

#### Outcome

According to this article, I can use my primary domain in the CSR and add additional domains (SANs) when activating the certificate.
