# PostgreSQL Replicator - Changelog

## Version 1.37 (2026-05-26)

**Fixed:**
- `None` handling in `get_all_tables_to_process()` - converts `None` to empty list when `tables:` section exists but is empty
- `None` handling in `copy_table()` - converts `None` to empty list when `nonvolatile:` section exists but is empty
- `None` handling in `generate_yaml_file()` - converts `None` to empty list for both `tables` and `nonvolatile` sections
- Error when `tables:` entry exists but is `None` (empty YAML section) causing `TypeError: 'NoneType' object is not iterable`

**Changed:**
- Auto-discovered tables are now included in generated YAML when `tables:` section is missing or `None`
- Existing tables configuration is preserved when `tables:` section has content
- Nonvolatile entries are preserved in generated YAML when they exist

## Version 1.36 (2026-05-26)

**Added:**
- `open_postgresql_connection_master()` method for network testing (connects to 'postgres' database)
- `row_exists_in_postgresql()` method to check if a row exists by its key columns
- `params` parameter to `pg_sql_execute()` for parameterized queries
- Counters for `duplicate_skipped_count` and `duplicate_rejected_count`
- Detailed logging for duplicate handling (distinguishes between skipped and rejected)
- Total elapsed runtime display at program completion (HH:MM:SS format)
- SLOW MODE now overrides nonvolatile optimization (copies tables even when row counts match)

**Changed:**
- Enhanced `insert_row()` to return tuple `(inserted, result_type, details)` with result types: `SUCCESS`, `DUPLICATE_SKIPPED`, `FK_VIOLATION`, `ERROR`
- Enhanced `print_validation_summary()` to write validation summary to log file (previously only console)
- Modified `_create_unique_constraint_on_base_table()` to skip automatic creation of unique constraints on child tables
- Modified `ensure_uniqueness_on_base_table()` to skip auto-creation with warning message
- `--slow` option no longer restricted to `--sync-deleted` only (can be used with normal replication)
- Non-volatile tables are now copied when `--slow` is enabled, regardless of row count match
- Added SLOW MODE diagnostic messages: "SLOW MODE: copying anyway (optimization disabled for debugging)"
- Added "SLOW MODE enabled - full sync" suffix to table sync messages
- Updated command line help text for `--slow` and `--nonvolatile` options

**Fixed:**
- Removed error condition that required `--slow` to be used with `--sync-deleted`

**Why:**
- Network testing now works even if target database doesn't exist (connects to default 'postgres' database)
- Duplicate handling now properly distinguishes between rows that exist (safe skip) and rows that were rejected (problem)
- Validation summary now persists in log file for historical reference
- PostgreSQL does not require child-side uniqueness for foreign keys, so auto-creation was unnecessary
- SLOW MODE can now be used for debugging nonvolatile optimization behavior

## Version 1.35 (2026-05-24)

**Added:**
- `nonvolatile` attribute to `ReplicationManager.__init__()`
- `--nonvolatile` command line flag
- `nonvolatile` section to YAML configuration
- Non-volatile table optimization in `copy_table()` method
- `nonvolatile` to `dump_internal_data()` for debugging

**Changed:**
- `copy_table()` now checks for non-volatile tables and skips copy when row counts match and `--nonvolatile` is enabled
- YAML generation now includes `nonvolatile` section

**Configuration example:**

nonvolatile:
  - BloodTestType
  - Insurance
  - Labs
  - RPR

## Version 1.34 (2026-05-24)

**Fixed:**
- "list index out of range" error in `_sync_deleted_table()` with bounds checking
- Validation reporting to re-query counts at display time

**Changed:**
- `print_validation_summary()` now re-queries current PostgreSQL counts instead of using cached values
- Added bounds checking when extracting key values for pagination in `_sync_deleted_table()`
- Added proper error handling for index extraction failures

**Why:**
- The "list index out of range" error occurred during sync-deleted batch processing when key indices didn't match row columns
- Validation summary sometimes showed stale counts due to cached values

## Version 1.33

**Baseline version** (existing code before changelog tracking)

**Core features at baseline:**
- MS Access to PostgreSQL replication
- DAO connection for Access (Windows-only via win32com)
- PostgreSQL connection via psycopg2
- Row-by-row processing (deliberate design choice)
- String concatenation for SQL (deliberate, due to parameterized query issues)
- Foreign key discovery from MS Access relationships
- Foreign key creation with ON DELETE CASCADE
- `--sync-deleted` for deletion synchronization
- `--slow` mode for sync-deleted (originally limited to sync-deleted only, expanded in v1.36)
- `--no-auto-index` to suppress automatic index/constraint creation
- `--full-refresh` to drop and recreate all tables
- `--schema` for schema-only replication
- Transformations (MMH3, yearonly, drop)
- Data validation and summary reporting
- Progress bars with ETA

---

## Command Line Options Summary (pg_replicator.py)

| Option | Description |
|--------|-------------|
| `-c, --config` | Path to configuration file (default: replicatorconfig.yaml) |
| `-s, --source` | MS Access database file name |
| `--thost` | PostgreSQL server host name or IP |
| `--tport` | PostgreSQL server port number |
| `--tdatabase` | PostgreSQL database name |
| `--tuser` | PostgreSQL user name |
| `--tpassword` | PostgreSQL password |
| `-v, --verbose` | Print informational messages |
| `--debug` | Enable SQL debugging output |
| `--trace` | Enable trace logging to file |
| `-a, --no-auto-index` | Suppress automatic creation of indexes/constraints for foreign keys |
| `--sync-deleted` | Synchronize deleted records from Access to PostgreSQL |
| `--slow` | Use slower processing; disables nonvolatile optimization (can be used with or without --sync-deleted) |
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

Version numbers follow a sequential revision scheme:
- **1.33** – Baseline
- **1.34** – Bug fixes (index error, validation reporting)
- **1.35** – New feature (nonvolatile optimization)
- **1.36** – Feature back-ports from ms_replicator.py (v1.1, v1.4, v1.5, v1.6, v1.8)
- **1.37** – Bug fixes (None handling in configuration sections)

## File Locations

| File | Description |
|------|-------------|
| `pg_replicator.py` | Main program (formerly replicator.py) |
| `pg_replicator.log` | Runtime log file |
| `replicatorconfig.yaml` | Configuration file |

---

## Notes

- This program targets **PostgreSQL** as the destination database
- A separate program `ms_replicator.py` exists for **Microsoft SQL Server** target
- Row-by-row processing and string concatenation for SQL are **deliberate design choices** (not oversights)
- The program is **Windows-only** due to DAO dependency for MS Access access

## Back-Ported Features from ms_replicator.py

| ms_replicator Version | Feature | Status in pg_replicator |
|----------------------|---------|------------------------|
| v1.1 | Master database connection for network testing | ✅ v1.36 |
| v1.4 | Enhanced result tracking from insert operations | ✅ v1.36 |
| v1.4 | Validation summary written to log file | ✅ v1.36 |
| v1.5 | Row existence checking method | ✅ v1.36 |
| v1.5 | Detailed logging for rejected rows | ✅ v1.36 |
| v1.5 | Distinguish between duplicate skipped and rejected | ✅ v1.36 |
| v1.6 | Skip automatic UNIQUE constraints on child tables | ✅ v1.36 |
| v1.8 | `--slow` option extended to normal replication | ✅ v1.36 |
| v1.8 | Total elapsed runtime display | ✅ v1.36 |
| v1.9 | None handling in configuration sections | ✅ v1.37 |

## Future Development

Remaining features from `ms_replicator.py` not yet back-ported:
- Data type size detection for TEXT fields (v1.2)
- Index compatibility checking (v1.2)
- NULL key handling in UPSERT (v1.3)
- Duplicate key detection with graceful handling (v1.3)
- Filtered unique index support (v1.7) - Less critical for PostgreSQL as it already handles NULLs correctly