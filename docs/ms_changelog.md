# MS SQL Server Replicator - Changelog

## Version 1.10 (2026-05-27)

**Added:**
- `scanned_tables` set to track which tables have been processed during sync-deleted
- `get_sqlserver_row_count()` helper method for reusable row count queries
- `get_parent_tables()` method to dynamically retrieve parent tables for a given child by scanning foreign keys
- Dependency-based iteration in `sync_deleted_tables()` for fast mode

**Changed:**
- `sync_deleted_tables()` now processes tables in dependency order (parents before children) when `--slow` is NOT enabled
- A table is only processed for deletion scanning when ALL its parent tables have already been scanned
- Tables with no parent dependencies are processed first
- Cascade deletions from parent tables can eliminate the need to scan child tables entirely
- Added diagnostic debug logging for tables waiting on parents

**Why:**
- Processing parents first allows cascade deletions to remove child rows before child tables are scanned
- Child tables whose row counts now match Access can be skipped entirely, reducing redundant work
- This optimization mirrors the dependency resolution used in `copy_all_fkey_tables()`
- Slow mode preserves original behavior for debugging and special cases

## Version 1.9 (2026-05-26)

**Fixed:**
- `None` handling in `get_all_tables_to_process()` - converts `None` to empty list when `tables:` section exists but is empty
- `None` handling in `copy_table()` - converts `None` to empty list when `nonvolatile:` section exists but is empty
- `None` handling in `generate_yaml_file()` - converts `None` to empty list for both `tables` and `nonvolatile` sections
- Error when `tables:` entry exists but is `None` (empty YAML section) causing `TypeError: 'NoneType' object is not iterable`

**Changed:**
- Auto-discovered tables are now included in generated YAML when `tables:` section is missing or `None`
- Existing tables configuration is preserved when `tables:` section has content
- Nonvolatile entries are preserved in generated YAML when they exist

## Version 1.8 (2026-05-26)

**Added:**
- Total elapsed runtime display at program completion (HH:MM:SS format, matching table-level output)
- `program_start_time` attribute to track overall execution time
- SLOW MODE now overrides nonvolatile optimization (copies tables even when row counts match)

**Changed:**
- `--slow` option no longer restricted to `--sync-deleted` only
- Non-volatile tables are now copied when `--slow` is enabled, regardless of row count match
- Added diagnostic messages for SLOW MODE: "SLOW MODE: copying anyway (optimization disabled for debugging)"
- Added "SLOW MODE enabled - full sync" suffix to table sync messages
- Updated command line help text for `--slow` and `--nonvolatile` options

**Fixed:**
- Removed error condition that required `--slow` to be used with `--sync-deleted`

## Version 1.7 (2026-05-26)

**Added:**
- Filtered unique index support for nullable columns
- Detection of nullable columns when creating UNIQUE indexes
- Automatic creation of filtered unique indexes (`WHERE column IS NOT NULL`) when all columns in a unique index are nullable
- Logging to indicate when filtered unique indexes are created

**Changed:**
- UNIQUE indexes on nullable columns now use filtered index syntax
- Multiple NULL values are now allowed in unique indexes (matches MS Access behavior)
- Uniqueness is still enforced for non-NULL values

**Why:**
- SQL Server's standard unique indexes reject multiple NULL values
- MS Access allows multiple NULLs in unique indexes
- This change bridges the behavioral difference and prevents replication failures

## Version 1.6 (2026-05-25)

**Fixed:**
- Indentation error in `copy_table()` method that caused syntax errors

**Changed:**
- Modified `_create_unique_constraint_on_base_table()` to no longer create unnecessary UNIQUE constraints on child tables
- Modified `ensure_uniqueness_on_base_table()` to skip automatic creation of unique constraints on child side of foreign keys
- Added warning logs when unique constraint creation is skipped
- Added documentation explaining that SQL Server does not require uniqueness on the child side of foreign keys

**Why:**
- SQL Server does not require the child side of a foreign key to be unique
- Auto-created UNIQUE constraints on nullable foreign key columns were causing replication failures (multiple NULLs rejected)
- This fix prevents those failures while maintaining referential integrity

## Version 1.5 (2026-05-25)

**Added:**
- `row_exists_in_sqlserver()` method to check if a row exists by its key columns
- Detailed logging for rejected rows including key values
- Counters for `duplicate_skipped_count` and `duplicate_rejected_count`

**Changed:**
- Enhanced `merge_row()` to return result_type: `SUCCESS`, `DUPLICATE_EXISTS`, `DUPLICATE_REJECTED`, `FK_VIOLATION`, `ERROR`
- When duplicate key error occurs, now checks if row actually exists in target
- Duplicate key error with existing row → `DUPLICATE_EXISTS` (safe skip)
- Duplicate key error with missing row → `DUPLICATE_REJECTED` (counted as failure, logged as warning)
- Completion message now shows both `duplicates skipped (exists)` and `duplicates rejected (missing)`
- Validation summary now written to log file (previously only console)

## Version 1.4 (2026-05-25)

**Added:**
- Enhanced `merge_row()` return tuple `(inserted, result_type, details)`
- Counters for `duplicate_skipped_count` and `duplicate_rejected_count`
- Validation summary now written to log file

**Changed:**
- Improved error tracking in `copy_table()`
- Completion message now shows duplicate counts and FK violations

## Version 1.3 (2026-05-25)

**Fixed:**
- MERGE statement syntax: added `source.` prefix for INSERT values
- NULL key handling: rows with NULL keys now use INSERT instead of MERGE
- Duplicate key detection (error 2601) now handled gracefully

**Changed:**
- `merge_row()` returns tuple with inserted flag and additional status flags

## Version 1.2 (2026-05-25)

**Fixed:**
- Data type mapping for `dbText` fields now uses `NVARCHAR(255)` or field-specific size instead of `NVARCHAR(MAX)`
- Index compatibility check added to skip indexes on `MAX` types (`NVARCHAR(MAX)`, `VARBINARY(MAX)`) and incompatible types (`TEXT`, `NTEXT`, `IMAGE`, `XML`)

**Added:**
- Warning logs when index creation is skipped due to incompatible column types
- Try-except in index creation to continue with remaining indexes if one fails

## Version 1.1 (2026-05-25)

**Added:**
- `open_sqlserver_connection_master()` method for network testing
- Connection to `master` database when `--network` flag is specified

**Changed:**
- Modified `test_network_connections()` to use `open_sqlserver_connection_master()` instead of `open_sqlserver_connection()`

## Version 1.0 (2026-05-25)

**Initial port from PostgreSQL replicator (pg_replicator.py v1.35)**

**Changes made:**
- Connection library: `psycopg2` → `pymssql`
- Configuration section: `postgresql:` → `sqlserver:`
- Command line options: `--thost`, `--tport`, etc. → `--shost`, `--sport`, etc.
- Identifier quoting: double quotes `"name"` → square brackets `[name]`
- Data type mappings for SQL Server compatibility:
  - `TEXT` → `NVARCHAR(MAX)`
  - `BYTEA` → `VARBINARY(MAX)`
  - `BOOLEAN` → `BIT`
  - `TIMESTAMP` → `DATETIME`
  - `SERIAL` → `IDENTITY`
  - `UUID` → `UNIQUEIDENTIFIER`
- UPSERT: `ON CONFLICT ... DO UPDATE` → `MERGE` statement
- Pagination: `LIMIT` with tuple comparison → `OFFSET ... FETCH NEXT`
- System catalog queries: `pg_*` tables → `sys.*` views
- `DROP TABLE CASCADE` → `DROP TABLE` without CASCADE (tables dropped in reverse dependency order)
- Schema: `public` → `dbo`
- Log file: `replicator.log` → `ms_replicator.log`
- Default config file: `replicatorconfig.yaml` → `ms_replicatorconfig.yaml`

**Features preserved:**
- `--nonvolatile` optimization
- `--sync-deleted` deletion synchronization
- `--slow` mode (originally only for sync-deleted, expanded in v1.8)
- `--no-auto-index` foreign key handling
- Transformations (`MMH3`, `yearonly`, `drop`)
- Foreign key discovery from MS Access relationships
- Validation and reporting
- Batch processing for deletions
- Progress bars with ETA

---

## Command Line Options Summary (ms_replicator.py)

| Option | Description |
|--------|-------------|
| `-c, --config` | Path to configuration file (default: ms_replicatorconfig.yaml) |
| `-s, --source` | MS Access database file name |
| `--shost` | SQL Server host name or IP |
| `--sport` | SQL Server port number |
| `--sdatabase` | SQL Server database name |
| `--suser` | SQL Server user name |
| `--spassword` | SQL Server password |
| `-v, --verbose` | Print informational messages |
| `--debug` | Enable SQL debugging output |
| `--trace` | Enable trace logging to file |
| `-a, --no-auto-index` | Suppress automatic creation of indexes/constraints for foreign keys |
| `--sync-deleted` | Synchronize deleted records from Access to SQL Server |
| `--slow` | Use slower processing; disables nonvolatile optimization and dependency resolution (can be used with or without --sync-deleted) |
| `--nonvolatile` | Skip copying non-volatile tables when row counts match (unless --slow is also enabled) |
| `-S, --schema` | Drop and recreate database, then replicate schema ONLY |
| `--adjust-ms-access` | Adjust MS Access schema (add AutoNumber primary key to tables without PK) |
| `-l, --list` | List table names and exit |
| `-n, --network` | Test both source and target connections |
| `--dump` | Dump internal program data |
| `--full-refresh` | Perform full refresh (drop and recreate all tables) |
| `-V, --version` | Show version and exit |
| `-o, --output` | Output file for generated YAML configuration |

---

## Version Format

Version numbers follow semantic versioning where practical:
- **Major** (1.x.x) – Significant changes, potential breaking changes
- **Minor** (x.1.x) – New features, backward compatible
- **Patch** (x.x.1) – Bug fixes, backward compatible

## File Locations

| File | Description |
|------|-------------|
| `ms_replicator.py` | Main program |
| `ms_replicator.log` | Runtime log file |
| `ms_replicatorconfig.yaml` | Configuration file |
| `ms_changelog.md` | This changelog |

## Version History Summary

| Version | Date | Focus |
|---------|------|-------|
| 1.0 | 2026-05-25 | Initial port from PostgreSQL |
| 1.1 | 2026-05-25 | Network testing with master database |
| 1.2 | 2026-05-25 | Data type fixes, index compatibility |
| 1.3 | 2026-05-25 | MERGE syntax, NULL key handling |
| 1.4 | 2026-05-25 | Enhanced result tracking, logging |
| 1.5 | 2026-05-25 | Row existence checking, duplicate handling |
| 1.6 | 2026-05-25 | Skip auto UNIQUE constraints on child tables |
| 1.7 | 2026-05-26 | Filtered unique indexes for nullable columns |
| 1.8 | 2026-05-26 | Extended --slow option, total elapsed time |
| 1.9 | 2026-05-26 | None handling in configuration sections |
| 1.10 | 2026-05-27 | Dependency-based sync-deleted processing (parents before children) |