# PostgreSQL Replicator - Changelog

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
- `--slow` mode for sync-deleted (processes all tables regardless of row counts)
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
| `--slow` | Use slower but safer deletion method (only valid with --sync-deleted) |
| `--nonvolatile` | Skip copying non-volatile tables when row counts match |
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

## File Locations

| File | Description |
|------|-------------|
| `pg_replicator.py` | Main program (formerly replicator.py) |
| `replicator.log` | Runtime log file |
| `replicatorconfig.yaml` | Configuration file |

---

## Notes

- This program targets **PostgreSQL** as the destination database
- A separate program `ms_replicator.py` exists for **Microsoft SQL Server** target
- Row-by-row processing and string concatenation for SQL are **deliberate design choices** (not oversights)
- The program is **Windows-only** due to DAO dependency for MS Access access

## Future Development

When ready, features from `ms_replicator.py` (v1.1 through v1.8) can be back-ported to `pg_replicator.py`:
- Filtered unique indexes (PostgreSQL already handles NULLs in unique indexes correctly)
- SLOW MODE override for nonvolatile optimization
- Enhanced duplicate handling with existence checks
- Total elapsed runtime display
- Improved validation logging to file