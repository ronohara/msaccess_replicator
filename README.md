# MS Access Replicator

A Python tool for replicating Microsoft Access databases to PostgreSQL
or MS SQL Server

## Description

This project has the objective of replicating any MS Access database
into the equivalent PostgreSQL database or MS SQL Seerver database
with as little configuration as possible.

In the AI research world, there is almost no development aimed at
integrating with user data stored in MS Access databases. There is
however significant support for PostgreSQL database, so by creating a
replicate of the user data, with the option to updated that replica at
will, it becomes possible to connect AI platforms to the data normally
being managed using MS Access.

This document describes the ‘replicator’ program – not the research
tools and processes that might use the replicated data.

There are two versions of the replicator. 

- pg_replicator.pg  - the target database is in a PostgrSQL server
- ms_replicator.py  - the target database in in an MS SQL server

The AI open context layer that we want to use the replicated database is
[WrenAI](https://github.com/Canner/WrenAI)

MS Access is a very effective UI and data management tool, but has
strong limitations related to capacity and robustness. Microsoft
recommends MS SQL Server for users where those limits become problem. A
consequence of that is that as of 2026, Microsoft has marked a number of
re-distributable support libraries for Access database as ‘no longer
supported’. That does not mean MS Access is unsupported, but given that
the UI component of Access can use MS SQL Server as its database, and MS
SQL Server has an ‘Express’ which has higher capacity limits than MS
Access and is free for use, you can deduce that the MS Access database
component is being strongly de-emphasized.

I would estimate that many many users of the original MS Access tool
have only modest transaction and storage capacity requirements, and very
limited budgets for software migrations, so a free tool like this allows
them to then use the amazing amount of development in the AI community for
analysing data.


## Where you would use the replicator

This replicator is a bare bones equivalent of MS SQL Server Migration
Assistant for a target PostgreSQL database or MS SQL database, which is then accessible
from many AI integration platforms.

However the intended use is not the same.

The replicator is just a normal program. But its functions allow you to easily maintain a replica of your data in a database system that is much better (directly) supported in a wide range of AI platforms.

We did experiment with Microsoft SQL Server Migration Assistant, but that is aimed at migrating away from MS Access rather than maintaining a replica - and it had lots of problems with some elements of the Access system we had as input – the same problems we had. Primarily with stored queries  It did not convert any of them, and we do not replicate them either.
Microsoft SQL Server Migration Assistant is an interactive migration assistance tool rather than a production process.

A big point is to leave your existing Access application(s) untouched. As a 'single source of truth' ... and the replication process becomes an automated workflow with zero manual steps to avoid that source of errors.

Lots of people are extracting Excel or CSV files and feeding them into a AI platform. That works of course, but is normally a manual approach - excellent for experiments, but not so good for regular operation.


## Installation

Refer to the installation notes in the MSAccess-pg-replicator or the MSAccess-ms-replicator PDF file in
the /docs directory

## Licensing note.

Although this is an MIT licensed project, the documentation is explicitly
excluded from that license.

It stands alone as a copyright document delivered in .pdf format and
is available for you to read under the usual fair use arrangement for
this sort of documentation.

We have done this so that this repository can be used as a central
point to coordinate revisions ... and retaining control of the
docummentation encourages people to work on this repository instead
of just forking their own version. But of course people remain free
to do that too.

