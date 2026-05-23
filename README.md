# MS PG Replicator

A Python tool for replicating Microsoft Access databases to PostgreSQL.

## Description

This project has the objective of replicating any MS Access database
into the equivalent PostgreSQL database with as little configuration as
possible.

In the AI research world, there is almost no development aimed at
integrating with user data stored in MS Access databases. There is
however significant support for PostgreSQL database, so by creating a
replicate of the user data, with the option to updated that replica at
will, it becomes possible to connect AI platforms to the data normally
being managed using MS Access.

This document describes the ‘replicator’ program – not the research
tools and processes that might use the replicated data.

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
them to then use the amazing amount development in the AI community for
analysing data.

This replicator is a bare bones equivalent of MS SQL Server Migration
Assistant for a target PostgreSQL database, which is then accessible
from many AI integration platforms.

## Installation

Refer to the installation notes in the MSAccess-replicator PDF file in
the /docs directory
