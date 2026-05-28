# MySQL Replicator - Changelog

## Version 1.1 (2026-05-28)

**Added:**
- `exit_program()` centralized shutdown method with connection cleanup and logging
- Log entry at program start recording version number
- Defensive exception handling in `close_connections()` for each connection type
- Connection handles set to None after closing in `close_connections()` and `test_network_connections()`

**Changed:**
- All `sys.exit(1)` calls replaced with `self.exit_program(1, error_msg)` for consistent shutdown handling
- `exit_and_cleanup()` simplified to call `exit_program(0)` instead of inline cleanup
- Main exception handler now uses `manager.exit_program()` when manager exists, with fallback for pre-manager errors
- PostgreSQL library (`psycopg2`) replaced with MySQL Connector (`mysql.connector`)
- Configuration section renamed from `postgresql:` to `mysql:`
- Command line options changed from `--thost`, `--tport`, etc. to `--mhost`, `--mport`, etc.
- Identifier quoting changed from double quotes to backticks
- `sanitise_token_for_postgresql()` renamed to `sanitise_token_for_mysql()` with backtick quoting
- `escape_postgresql_string()` renamed to `escape_mysql_string()`
- `convert_dao_value_to_postgresql_literal()` renamed to `convert_dao_value_to_mysql_literal()`
- `get_postgresql_row_count()` renamed to `get_mysql_row_count()`
- `table_exists_in_postgresql()` renamed to `table_exists_in_mysql()` using MySQL information_schema
- `row_exists_in_postgresql()` renamed to `row_exists_in_mysql()`
- `pg_sql_execute()` renamed to `mysql_sql_execute()` with `params` parameter removed
- `open_postgresql_connection()` renamed to `open_mysql_connection()`
- `open_postgresql_connection_master()` renamed to `open_mysql_connection_master()`
- Data type mappings converted from PostgreSQL types to MySQL types
- UPSERT syntax changed from `ON CONFLICT ... DO UPDATE` to `ON DUPLICATE KEY UPDATE`
- Pagination in `_sync_deleted_table()` changed from tuple comparison to `LIMIT ... OFFSET`
- System catalog queries converted from PostgreSQL `pg_*` tables to MySQL information_schema
- Primary key and unique constraint checks now use information_schema
- Foreign key existence checks now use information_schema
- Schema replication now uses MySQL `CREATE DATABASE` with utf8mb4 character set
- Table creation now includes `ENGINE=InnoDB` for foreign key support

**Fixed:**
- Connection cleanup now occurs consistently on all exit paths
- Error messages now properly logged before program termination in all scenarios
- DAO connection warning "Object invalid or no longer set" now handled gracefully (logged at debug level)
- Connection handles now properly set to None after closing to prevent double-close attempts

**Why:**
- Centralized exit logic ensures connections are always closed properly regardless of how the program terminates
- Consistent logging of termination messages (success or error) for audit trail
- Prevents resource leaks from early exits that bypassed connection cleanup
- Provides unified error reporting format across all failure points
- MySQL requires different UPSERT syntax (`ON DUPLICATE KEY UPDATE` instead of `ON CONFLICT`)
- MySQL does not support tuple comparison for pagination, requiring OFFSET approach
- MySQL uses backticks for quoted identifiers instead of double quotes
- MySQL Connector library provides native MySQL support

---

## Version 1.0 (2026-05-28)

**Initial port from PostgreSQL replicator (pg_replicator.py v1.39)**

Changes made from pg_replicator v1.39:
- Connection library: psycopg2 → mysql.connector
- Configuration section: postgresql: → mysql:
- Command line options: --thost, --tport, etc. → --mhost, --mport, etc.
- Identifier quoting: double quotes "name" → backticks `name`
- Data type mappings for MySQL compatibility (BOOLEAN, INT, BIGINT, DATETIME, BLOB, TEXT, etc.)
- UPSERT: ON CONFLICT ... DO UPDATE → ON DUPLICATE KEY UPDATE
- Pagination: tuple comparison (col) > (val) → LIMIT ... OFFSET
- System catalog queries: pg_* tables → information_schema views
- Schema: public → database() function
- DROP TABLE CASCADE → DROP TABLE (no CASCADE needed)
- Table engine: InnoDB for foreign key support
- Character set: utf8mb4 with utf8mb4_unicode_ci collation

Features preserved:
- --nonvolatile optimization
- --sync-deleted deletion synchronization
- --slow mode (affects nonvolatile optimization and dependency resolution)
- --no-auto-index foreign key handling
- Transformations (MMH3, yearonly, drop)
- Foreign key discovery from MS Access relationships
- Validation and reporting
- Batch processing for deletions (OFFSET pagination)
- Progress bars with ETA
- Dependency-based sync-deleted processing (parents before children)

---

## Command Line Options Summary

| Option | Description |
|--------|-------------|
| `-c, --config` | Path to configuration file (default: my_replicatorconfig.yaml) |
| `-s, --source` | MS Access database file name |
| `--mhost` | MySQL server host name or IP |
| `--mport` | MySQL server port number |
| `--mdatabase` | MySQL database name |
| `--muser` | MySQL user name |
| `--mpassword` | MySQL password |
| `-v, --verbose` | Print informational messages |
| `--debug` | Enable SQL debugging output |
| `--trace` | Enable trace logging to file |
| `-a, --no-auto-index` | Suppress automatic creation of indexes/constraints for foreign keys |
| `--sync-deleted` | Synchronize deleted records from Access to MySQL |
| `--slow` | Use slower processing; disables nonvolatile optimization and dependency resolution |
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
- Major (1.x.x) – Significant changes, potential breaking changes
- Minor (x.1.x) – New features, backward compatible
- Patch (x.x.1) – Bug fixes, backward compatible

## File Locations

| File | Description |
|------|-------------|
| my_replicator.py | Main program |
| my_replicator.log | Runtime log file |
| my_replicatorconfig.yaml | Configuration file |

## Version History Summary

| Version | Date | Focus |
|---------|------|-------|
| 1.0 | 2026-05-28 | Initial port from PostgreSQL replicator |
| 1.1 | 2026-05-28 | Centralized exit handling, string concatenation for SQL, defensive connection cleanup, OFFSET pagination, ON DUPLICATE KEY UPDATE |