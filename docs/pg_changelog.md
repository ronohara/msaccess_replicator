# PostgreSQL Replicator - Changelog

## Version 1.40 (2026-05-31)

### Added
- `--create-views` command line option to create sane views on an existing PostgreSQL database without schema changes or data copy
- `--simple-names` command line option to create tables and columns with simple lowercase names (no quoted identifiers, spaces replaced with underscores)
- `--schema` and `--create-views` can now be used together (replaces the removed `--schema-and-views`)
- `internal_replicator_data` metadata table to record whether `--simple-names` mode was used during schema creation
- `create_internal_replicator_table()` method to create the metadata table
- `read_internal_replicator_data()` method to read the `simplenames` setting from the database
- `get_sanitise_function()` method that returns either `sanitise_token_for_postgresql` (quoted mode) or `sanitise_for_sane_view` (simple names mode) based on the `simple_names` flag
- Special character handling in `sanitise_for_sane_view()`:
  - `%` → `_percent`
  - `$` → `dollar_`, `_dollar_`, or `_dollar` based on position
  - `#` → `hash_`, `_hash_`, or `_hash` based on position
  - `@` → `at_`, `_at_`, or `_at` based on position
  - `&` → `amp_`, `_amp_`, or `_amp` based on position
  - `*` → `star_`, `_star_`, or `_star` based on position
  - `+` → `plus_`, `_plus_`, or `_plus` based on position
  - `-` → `minus_`, `_minus_`, or `_minus` based on position
  - `(` → `lbrk_`, `_lbrk_`, or `_lbrk` based on position
  - `)` → `rbrk_`, `_rbrk_`, or `_rbrk` based on position
- Leading digit handling in `sanitise_for_sane_view()` (prefixes with underscore)

### Changed
- `create_views_only()` now reads `simplenames` from `internal_replicator_data` to determine naming mode for table lookups
- `create_views_only()` now opens both MS Access and PostgreSQL connections (was PostgreSQL-only)
- `create_sane_views()` and `create_views_only()` now use `col['name']` (the sanitised column identifier) in the SELECT clause instead of quoting the original column name
- `check_primary_key()`, `check_unique_constraint_only()`, and `check_reference_table_has_uniqueness()` now sanitise column names before comparing with PostgreSQL schema (fixes foreign key detection when `--simple-names` is active)
- `create_foreign_key()` now correctly handles the case where `ensure_uniqueness_on_base_table()` returns `None`, logging an error and skipping instead of entering an infinite recursive loop
- `create_all_indexes()` now uses the sanitised table name for existence checks instead of the original Access table name
- `create_foreign_key()` existence check now uses the sanitised table name
- `load_table()` now uses `get_sanitise_function()` to determine the appropriate sanitisation method
- All SQL generation now respects the `simple_names` flag via the dynamic sanitisation function

### Removed
- `--schema-and-views` command line option (replaced by using `--schema` and `--create-views` together)
- `schema_and_views` flag from `ReplicationManager`
- All `column_mapping` references (undocumented feature that was never requested)

### Fixed
- Foreign key infinite loop when neither referenced table has a PRIMARY KEY or UNIQUE constraint
- Column name collisions in simple-names mode (e.g., `nadir` and `nadir%` now become `nadir` and `nadir_percent`)
- Identifiers starting with digits in simple-names mode now correctly prefixed with underscore
- `create_views_only()` now uses the correct table name when the database was created with `--simple-names`
- Index creation existence check now correctly identifies existing indexes when `--simple-names` is active
- Foreign key existence check now correctly identifies existing foreign keys when `--simple-names` is active
- Trailing underscore stripping in `sanitise_for_sane_view` no longer removes meaningful underscores (e.g., `nadir_` from `nadir%` is preserved)

## Version 1.39 (2026-05-27)

### Added
- `exit_program()` centralized shutdown method with connection cleanup and logging
- Log entry at program start recording version number

### Changed
- All `sys.exit(1)` calls replaced with `self.exit_program(1, error_msg)` for consistent shutdown handling
- `exit_and_cleanup()` simplified to call `exit_program(0)` instead of inline cleanup
- Main exception handler now uses `manager.exit_program()` when manager exists, with fallback for pre-manager errors

### Fixed
- Connection cleanup now occurs consistently on all exit paths
- Error messages now properly logged before program termination in all scenarios

## Version 1.38 (2026-05-27)

### Added
- `scanned_tables` set to track which tables have been processed during sync-deleted
- `get_postgresql_row_count()` helper method for reusable row count queries
- `get_parent_tables()` method to dynamically retrieve parent tables for a given child by scanning foreign keys
- Dependency-based iteration in `sync_deleted_tables()` for fast mode

### Changed
- `sync_deleted_tables()` now processes tables in dependency order (parents before children) when `--slow` is NOT enabled
- A table is only processed for deletion scanning when ALL its parent tables have already been scanned
- Tables with no parent dependencies are processed first
- Cascade deletions from parent tables can eliminate the need to scan child tables entirely
- Added diagnostic debug logging for tables waiting on parents

## Version 1.37 (2026-05-26)

### Fixed
- `None` handling in `get_all_tables_to_process()` - converts `None` to empty list when `tables:` section exists but is empty
- `None` handling in `copy_table()` - converts `None` to empty list when `nonvolatile:` section exists but is empty
- `None` handling in `generate_yaml_file()` - converts `None` to empty list for both `tables` and `nonvolatile` sections
- Error when `tables:` entry exists but is `None` (empty YAML section) causing type error

### Changed
- Auto-discovered tables are now included in generated YAML when `tables:` section is missing or `None`
- Existing tables configuration is preserved when `tables:` section has content
- Nonvolatile entries are preserved in generated YAML when they exist

## Version 1.36 (2026-05-26)

### Added
- `open_postgresql_connection_master()` method for network testing (connects to 'postgres' database)
- `row_exists_in_postgresql()` method to check if a row exists by its key columns
- Counters for duplicate skipped and duplicate rejected
- Detailed logging for duplicate handling distinguishing between skipped and rejected
- Total elapsed runtime display at program completion
- Slow mode now overrides nonvolatile optimization (copies tables even when row counts match)

### Changed
- Enhanced `insert_row()` to return tuple with result types: SUCCESS, DUPLICATE_SKIPPED, FK_VIOLATION, ERROR
- Enhanced `print_validation_summary()` to write validation summary to log file (previously only console)
- Modified `_create_unique_constraint_on_base_table()` to skip automatic creation of unique constraints on child tables
- Modified `ensure_uniqueness_on_base_table()` to skip auto-creation with warning message
- `--slow` option no longer restricted to `--sync-deleted` only (can be used with normal replication)
- Non-volatile tables are now copied when `--slow` is enabled, regardless of row count match
- Added slow mode diagnostic messages
- Updated command line help text for `--slow` and `--nonvolatile` options

### Fixed
- Removed error condition that required `--slow` to be used with `--sync-deleted`

## Version 1.35 (2026-05-24)

### Added
- `nonvolatile` attribute to ReplicationManager
- `--nonvolatile` command line flag
- `nonvolatile` section to YAML configuration
- Non-volatile table optimization in copy_table method
- `nonvolatile` to dump_internal_data for debugging

### Changed
- copy_table now checks for non-volatile tables and skips copy when row counts match and `--nonvolatile` is enabled
- YAML generation now includes nonvolatile section

## Version 1.34 (2026-05-24)

### Fixed
- List index out of range error in sync_deleted_table with bounds checking
- Validation reporting to re-query counts at display time

### Changed
- print_validation_summary now re-queries current PostgreSQL counts instead of using cached values
- Added bounds checking when extracting key values for pagination in sync_deleted_table
- Added proper error handling for index extraction failures

## Version 1.33

### Baseline Version

Core features at baseline:
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

## Command Line Options Summary

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
| `--slow` | Use slower processing; disables nonvolatile optimization and dependency resolution |
| `--nonvolatile` | Skip copying non-volatile tables when row counts match (unless --slow is also enabled) |
| `-S, --schema` | Drop and recreate database, then replicate schema ONLY |
| `--create-views` | Create sane views on existing database (no schema changes, no data copy). When used with --schema, creates schema THEN views. |
| `--simple-names` | Create tables and columns with simple lowercase names (no quoted identifiers). Can only be used with --schema. |
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
- 1.33 – Baseline
- 1.34 – Bug fixes (index error, validation reporting)
- 1.35 – New feature (nonvolatile optimization)
- 1.36 – Feature back-ports and enhancements
- 1.37 – Bug fixes (None handling in configuration sections)
- 1.38 – Dependency-based sync-deleted processing (parents before children)
- 1.39 – Centralized exit handling and shutdown cleanup
- 1.40 – Simple names mode, view creation options, special character handling, and foreign key fixes

## File Locations

| File | Description |
|------|-------------|
| pg_replicator.py | Main program |
| pg_replicator.log | Runtime log file |
| replicatorconfig.yaml | Configuration file |

---

## Notes

- This program targets PostgreSQL as the destination database
- Row-by-row processing and string concatenation for SQL are deliberate design choices
- The program is Windows-only due to DAO dependency for MS Access access
- When `--simple-names` is used during schema creation, subsequent replication runs automatically detect and use simple naming mode via the `internal_replicator_data` table
- Special character replacements in simple-names mode use descriptive words (e.g., `%` becomes `_percent`, `$` becomes `_dollar_`, `(` becomes `_lbrk_`, `)` becomes `_rbrk_`) to avoid collisions and maintain readability