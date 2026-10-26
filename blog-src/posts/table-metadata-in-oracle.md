---
title: "Table metadata in Oracle"
date: 2020-06-15
slug: table-metadata-in-oracle
---

#### Context

Normally I work in SQL Server when I am dealing with relational data but on my current project one of our primary systems uses Oracle as its back end. In order to quickly generate a list of column names for use in an access request, I need to interrogate the table metadata and I cannot recall the way to do this on Oracle.

#### Search terms

oracle column list query

#### Helpful link

https://dataedo.com/kb/query/oracle/list-columns-names-in-specific-table

#### Outcome

I was reminded to use sys.all_tab_columns to get this information.
