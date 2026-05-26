#!/usr/bin/env python
# -*- coding: utf-8 -*-

__version__ = "$Revision: 1.37 $"

import sys
import os
import re
import yaml
import argparse
import hashlib
import mmh3
import inspect
import traceback
import subprocess
from datetime import datetime, date
from decimal import Decimal
from typing import List, Dict, Any, Optional
import logging
import logging.handlers
import psycopg2
import psycopg2.extras
import win32com.client
import time

# ============================================================================
# LOGGING CONFIGURATION
# ============================================================================

logger = logging.getLogger('pg_replicator')
logger.setLevel(logging.DEBUG)

file_handler = logging.handlers.RotatingFileHandler(
    'pg_replicator.log',
    maxBytes=10 * 1024 * 1024,
    backupCount=10,
    encoding='utf-8'
)
file_handler.setLevel(logging.DEBUG)

formatter = logging.Formatter('{asctime} - {levelname} - {message}', datefmt='%Y-%m-%d %H:%M', style='{')
file_handler.setFormatter(formatter)

logger.addHandler(file_handler)

console_handler = logging.StreamHandler()
console_handler.setFormatter(formatter)
console_handler.setLevel(logging.INFO)
logger.addHandler(console_handler)

# ============================================================================
# DECODE FUNCTION AND DATA TYPE CONVERSION
# ============================================================================

def decode_sketchy_utf16(raw_bytes):
    if raw_bytes is None:
        return ''
    if isinstance(raw_bytes, str):
        return raw_bytes
    if isinstance(raw_bytes, bytes):
        try:
            return raw_bytes.decode('utf-16', errors='replace')
        except:
            try:
                return raw_bytes.decode('utf-8', errors='replace')
            except:
                return str(raw_bytes)
    return str(raw_bytes)

def escape_postgresql_string(value):
    """Escape a string for safe concatenation into PostgreSQL SQL (without quotes)"""
    if value is None:
        return 'NULL'
    if isinstance(value, str):
        # Replace single quote with two single quotes
        escaped = value.replace("'", "''")
        # Replace backslash with double backslash
        escaped = escaped.replace('\\', '\\\\')
        return escaped
    return str(value)

def convert_dao_value_to_postgresql_literal(value):
    """Convert DAO value to PostgreSQL literal string for SQL concatenation"""
    if value is None:
        return 'NULL'
    elif isinstance(value, bool):
        return 'TRUE' if value else 'FALSE'
    elif isinstance(value, str):
        return f"'{escape_postgresql_string(value)}'"
    elif isinstance(value, (int, float, Decimal)):
        return str(value)
    elif isinstance(value, datetime):
        return f"'{value.strftime('%Y-%m-%d %H:%M:%S')}'"
    elif isinstance(value, date):
        return f"'{value.strftime('%Y-%m-%d')}'"
    else:
        return f"'{escape_postgresql_string(str(value))}'"

def convert_dao_value_to_python(value):
    if value is None:
        return None
    if hasattr(value, 'year') and hasattr(value, 'month') and hasattr(value, 'day'):
        if hasattr(value, 'hour') and hasattr(value, 'minute') and hasattr(value, 'second'):
            try:
                return datetime(value.year, value.month, value.day, value.hour, value.minute, value.second)
            except:
                return date(value.year, value.month, value.day)
        else:
            try:
                return date(value.year, value.month, value.day)
            except:
                return str(value)
    if isinstance(value, Decimal):
        return value
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value
    return str(value)

# ============================================================================
# MS ACCESS RESERVED WORDS (partial list)
# ============================================================================

ms_access_reserved = {
    'SELECT', 'INSERT', 'UPDATE', 'DELETE', 'FROM', 'WHERE', 'GROUP', 'BY',
    'ORDER', 'HAVING', 'JOIN', 'INNER', 'LEFT', 'RIGHT', 'ON', 'AS', 'AND',
    'OR', 'NOT', 'NULL', 'IS', 'LIKE', 'IN', 'BETWEEN', 'EXISTS', 'ALL',
    'ANY', 'DISTINCT', 'TOP', 'PERCENT', 'UNION', 'ALL', 'INTO', 'VALUES',
    'SET', 'CREATE', 'ALTER', 'DROP', 'TABLE', 'INDEX', 'VIEW', 'PROCEDURE',
    'FUNCTION', 'TRIGGER', 'CONSTRAINT', 'PRIMARY', 'KEY', 'FOREIGN', 'REFERENCES',
    'UNIQUE', 'CHECK', 'DEFAULT', 'COLUMN', 'DATABASE', 'SCHEMA', 'GRANT',
    'REVOKE', 'TRANSACTION', 'COMMIT', 'ROLLBACK', 'SAVEPOINT', 'BEGIN', 'END'
}

# ============================================================================
# REPLICATION MANAGER CLASS
# ============================================================================

class ReplicationManager:
    
    def __init__(self):
        self.trace = False
        self.verbose = False
        self.debug = False
        self.list_tables = False
        self.test_network = False
        self.dump_data = False
        self.schema_only = False
        self.current_fk_iteration = 0
        self.no_auto_index = False          # suppress automatic index/constraint creation
        self.adjust_ms_access = False       # adjust MS Access schema flag
        self.full_refresh = False           # full refresh flag
        self.sync_deleted = False           # sync deleted records flag (boolean)
        self.slow = False                   # slow deletion mode flag (also affects nonvolatile optimization)
        self.nonvolatile = False            # enable non-volatile table optimization
        
        self.parameters = {}
        self.dao_conn = None
        self.pg_conn = None
        self.pg_cursor = None
        self.original_configured_fkeys = {}
        self.discovered_foreign_keys = {}
        self.copied_tables = set()
        
        # Progress tracking
        self.last_progress_time = 0
        self.progress_interval = 2
        
        # Data validation tracking
        self.validation_results = {}
        
        # Track empty tables (0 rows)
        self.empty_tables = []
        
        # Program start time for total elapsed calculation
        self.program_start_time = None
        
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - __init__ called")
    
    # ========================================================================
    # SYSTEM TABLE IDENTIFICATION
    # ========================================================================
    
    def is_system_table_name(self, table_name):
        """Check if a table name indicates it's a system table."""
        if not table_name:
            return True
        name_lower = table_name.lower()
        # MS Access system tables start with MSys
        if name_lower.startswith('msys'):
            return True
        # Temporary tables start with ~
        if table_name.startswith('~'):
            return True
        # Add any other patterns you want to treat as system tables
        return False
    
    # ========================================================================
    # FORMATTING HELPERS
    # ========================================================================
    
    def format_eta(self, seconds):
        """Convert seconds to HH:MM:SS format for ETA display."""
        if seconds < 0:
            seconds = 0
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        secs = int(seconds % 60)
        
        if hours > 0:
            return f"{hours:02d}:{minutes:02d}:{secs:02d}"
        elif minutes > 0:
            return f"{minutes:02d}:{secs:02d}"
        else:
            return f"{secs}s"
    
    def format_elapsed(self, seconds):
        """Convert seconds to HH:MM:SS format for elapsed time display."""
        if seconds < 0:
            seconds = 0
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        secs = int(seconds % 60)
        
        if hours > 0:
            return f"{hours:02d}:{minutes:02d}:{secs:02d}"
        elif minutes > 0:
            return f"{minutes:02d}:{secs:02d}"
        else:
            return f"{secs:02d}s"
    
    # ========================================================================
    # CONFIGURATION VALIDATION
    # ========================================================================
    
    def validate_configuration(self):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - validate_configuration called")
        
        logger.info("Validating configuration...")
        
        pg_params = self.parameters.get('postgresql', {})
        required_pg = ['host', 'port', 'database', 'user', 'password']
        missing_pg = [p for p in required_pg if not pg_params.get(p)]
        
        if missing_pg:
            error_msg = f"Missing PostgreSQL parameters: {', '.join(missing_pg)}"
            logger.error(error_msg)
            raise ValueError(error_msg)
        
        try:
            pg_params['port'] = int(pg_params['port'])
        except (ValueError, TypeError):
            error_msg = f"PostgreSQL port must be an integer, got: {pg_params.get('port')}"
            logger.error(error_msg)
            raise ValueError(error_msg)
        
        dao_params = self.parameters.get('DAO', {})
        if not dao_params.get('database'):
            error_msg = "Missing MS Access database path in configuration"
            logger.error(error_msg)
            raise ValueError(error_msg)
        
        db_path = dao_params.get('database')
        if not os.path.exists(db_path) and not os.path.isabs(db_path):
            db_path = os.path.abspath(db_path)
        if not os.path.exists(db_path):
            error_msg = f"MS Access database file not found: {db_path}"
            logger.error(error_msg)
            raise ValueError(error_msg)
        
        logger.info("Configuration validation passed")
        if self.verbose:
            print("Configuration validation passed")
        
        return True
    
    # ========================================================================
    # VALIDATE TRANSFORMATIONS
    # ========================================================================
    
    def validate_transformation(self):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - validate_transformation called")
        
        transformations = self.parameters.get('transformations', [])
        
        if not isinstance(transformations, list):
            logger.info("No transformations defined in configuration")
            return True
        
        if len(transformations) == 0:
            logger.info("No transformations defined in configuration")
            return True
        
        logger.info("Validating transformations against Access database schema...")
        
        # Build a set of existing table names for quick lookup
        existing_tables = set()
        for i in range(self.dao_conn.TableDefs.Count):
            tdef = self.dao_conn.TableDefs[i]
            if not self.is_system_table(tdef):
                existing_tables.add(decode_sketchy_utf16(tdef.Name))
        
        # Build a dictionary of existing column names per table
        existing_columns = {}
        for table_name in existing_tables:
            tdef = self.dao_conn.TableDefs[table_name]
            columns = set()
            for i in range(tdef.Fields.Count):
                fld = tdef.Fields[i]
                columns.add(decode_sketchy_utf16(fld.Name))
            existing_columns[table_name] = columns
        
        valid = True
        warning_count = 0
        
        for trans in transformations:
            if not isinstance(trans, dict):
                logger.warning(f"Skipping invalid transformation entry (not a dict): {trans}")
                continue
            
            trans_table = trans.get('table', '')
            if not trans_table:
                logger.warning(f"Skipping transformation with missing 'table' field: {trans}")
                continue
            
            # Check if table exists in Access
            if trans_table not in existing_tables:
                logger.warning(f"Transformation table '{trans_table}' not found in Access database - transformations for this table will NOT be applied")
                valid = False
                warning_count += 1
                continue
            
            columns_list = trans.get('columns', [])
            if not isinstance(columns_list, list):
                logger.warning(f"Transformation for table '{trans_table}' has invalid 'columns' field (not a list) - skipping")
                continue
            
            # Check each column in the transformation
            for col_entry in columns_list:
                if not isinstance(col_entry, dict):
                    logger.warning(f"Transformation for table '{trans_table}' has invalid column entry (not a dict): {col_entry}")
                    continue
                
                col_name = col_entry.get('name', '')
                if not col_name:
                    logger.warning(f"Transformation for table '{trans_table}' has column entry with missing 'name' field: {col_entry}")
                    continue
                
                transform_type = col_entry.get('transform', '')
                if not transform_type:
                    logger.warning(f"Transformation for table '{trans_table}' column '{col_name}' has missing 'transform' field - will be ignored")
                    continue
                
                # Check if column exists in the table
                if col_name not in existing_columns.get(trans_table, set()):
                    logger.warning(f"Transformation column '{col_name}' not found in table '{trans_table}' - this transformation will NOT be applied")
                    valid = False
                    warning_count += 1
        
        if valid:
            logger.info("All transformations validated successfully against Access database schema")
        else:
            logger.warning(f"Transformation validation completed with {warning_count} warning(s) - affected transformations will be skipped")
        
        return valid
    
    # ========================================================================
    # PROGRESS INDICATION
    # ========================================================================
    
    def show_progress(self, table_name, current, total, stage="Copying", start_time=None):
        """Show progress bar with ETA calculation based on start time."""
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - show_progress called with params: table_name={table_name}, current={current}, total={total}, stage={stage}")
        
        current_time = time.time()
        
        # Only update display at intervals or when complete
        if current_time - self.last_progress_time >= self.progress_interval or current == total:
            percentage = (current / total * 100) if total > 0 else 0
            
            # Build progress bar
            bar_length = 40
            filled_length = int(bar_length * current // total) if total > 0 else 0
            bar = '█' * filled_length + '░' * (bar_length - filled_length)
            
            # Calculate ETA using start_time if provided
            eta_str = ""
            if start_time is not None and current > 0 and current < total:
                elapsed_since_start = current_time - start_time
                if elapsed_since_start > 0:
                    rate = current / elapsed_since_start
                    remaining_seconds = (total - current) / rate
                    eta_str = f" ETA: {self.format_eta(remaining_seconds)}"
            
            log_message = f"{stage} {table_name}: [{bar}] {percentage:.1f}% ({current}/{total}){eta_str}"
            
            if self.verbose:
                if current < total:
                    print(f"\r{log_message}", end='', flush=True)
                else:
                    print(f"\r{log_message}")
            else:
                logger.debug(log_message)
            
            self.last_progress_time = current_time
    
    # ========================================================================
    # GET ACCURATE ROW COUNT FROM ACCESS USING SQL
    # ========================================================================
    
    def get_access_row_count(self, table_name):
        """Get accurate row count from Access using SQL COUNT"""
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - get_access_row_count called for {table_name}")
        
        try:
            # Use SQL COUNT for accurate row count
            count_sql = f"SELECT COUNT(*) FROM [{table_name}]"
            recordset = self.dao_conn.OpenRecordset(count_sql)
            count = recordset.Fields(0).value
            recordset.Close()
            return count
        except Exception as e:
            logger.error(f"Failed to get row count for {table_name}: {e}")
            # Fallback to older method
            tdef = self.dao_conn.TableDefs[table_name]
            recordset = tdef.OpenRecordset()
            if recordset.EOF and recordset.BOF:
                return 0
            else:
                recordset.MoveLast()
                count = recordset.RecordCount
                recordset.Close()
                return count
    
    # ========================================================================
    # DATA VALIDATION
    # ========================================================================
    
    def validate_table_data(self, table_name, safe_table_name):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - validate_table_data called with params: table_name={table_name}, safe_table_name={safe_table_name}")
        
        logger.info(f"Validating data for table: {table_name}")
        
        try:
            # Use SQL COUNT for accurate source count
            source_count = self.get_access_row_count(table_name)
            
            count_sql = f"SELECT COUNT(*) FROM {safe_table_name}"
            result = self.pg_sql_execute(count_sql, fetch_one=True)
            target_count = result[0] if result else 0
            
            self.validation_results[table_name] = {
                'source_count': source_count,
                'target_count': target_count,
                'matched': source_count == target_count,
                'difference': source_count - target_count
            }
            
            if source_count == target_count:
                msg = f"✓ VALIDATED: {table_name} (Access: {source_count}, PostgreSQL: {target_count})"
                logger.info(msg)
                print(msg)
                return True
            else:
                msg = f"✗ VALIDATION FAILED: {table_name} - Access: {source_count}, PostgreSQL: {target_count}, Difference: {source_count - target_count}"
                logger.warning(msg)
                print(msg)
                return False
                
        except Exception as e:
            logger.error(f"Validation error for {table_name}: {e}")
            self.validation_results[table_name] = {
                'source_count': None,
                'target_count': None,
                'matched': False,
                'error': str(e)
            }
            return False
    
    def print_validation_summary(self):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - print_validation_summary called")
        
        # Build summary lines for both console and log
        summary_lines = []
        
        summary_lines.append("\n" + "=" * 110)
        summary_lines.append("DATA VALIDATION SUMMARY")
        summary_lines.append("=" * 110)
        
        # Re-query current counts for accurate display
        fresh_counts = {}
        for table_name in self.validation_results.keys():
            try:
                safe_table_name = self.sanitise_token_for_postgresql(table_name)
                count_sql = f"SELECT COUNT(*) FROM {safe_table_name}"
                result = self.pg_sql_execute(count_sql, fetch_one=True)
                fresh_counts[table_name] = result[0] if result else 0
            except Exception as e:
                logger.warning(f"Could not re-query count for {table_name}: {e}")
                fresh_counts[table_name] = self.validation_results[table_name].get('target_count', 'N/A')
        
        # Check if any sync_deleted operation was performed
        sync_was_performed = any('pg_count_before' in result for result in self.validation_results.values())
        
        if sync_was_performed:
            # Extended display with before/after counts and synced flag
            summary_lines.append("\nPer-table validation results (after sync_deleted):")
            summary_lines.append("-" * 110)
            summary_lines.append(f"{'Table Name':<30} {'Access':>10} {'PG (Before)':>12} {'PG (After)':>11} {'Deleted':>10} {'Status':>8} {'Synced':>6}")
            summary_lines.append("-" * 110)
            
            for table_name, result in self.validation_results.items():
                source_count = result.get('source_count', 'N/A')
                # Use fresh count for PG After
                pg_after = fresh_counts.get(table_name, result.get('target_count', 'N/A'))
                deleted_count = result.get('deleted', 0)
                was_synced = 'pg_count_before' in result
                
                if was_synced:
                    pg_before = result.get('pg_count_before', 'N/A')
                    expected_deleted = pg_before - pg_after if isinstance(pg_before, int) and isinstance(pg_after, int) else 0
                    status = "✓ PASS" if (source_count == pg_after and deleted_count == expected_deleted) else "✗ FAIL"
                    synced_flag = "Yes"
                else:
                    pg_before = 'N/A'
                    status = "✓ PASS" if source_count == pg_after else "✗ FAIL"
                    synced_flag = "No"
                
                display_name = table_name[:29] if len(table_name) > 29 else table_name
                deleted_display = str(deleted_count) if deleted_count > 0 else ""
                
                summary_lines.append(f"{display_name:<30} {source_count:>10} {pg_before:>12} {pg_after:>11} {deleted_display:>10} {status:>8} {synced_flag:>6}")
            
            summary_lines.append("-" * 110)
            
            # Additional consistency check summary - ONLY for tables that were synced
            inconsistent_tables = []
            for table_name, result in self.validation_results.items():
                # Only check tables that were processed by sync_deleted (have pg_count_before)
                if 'pg_count_before' not in result:
                    continue  # Skip tables that were never synced
                
                source_count = result.get('source_count', 0)
                pg_after = fresh_counts.get(table_name, result.get('target_count', 0))
                pg_before = result.get('pg_count_before', 0)
                deleted = result.get('deleted', 0)
                
                if source_count != pg_after:
                    inconsistent_tables.append((table_name, 'count_mismatch', source_count, pg_after))
                elif deleted != (pg_before - pg_after):
                    inconsistent_tables.append((table_name, 'deleted_mismatch', deleted, pg_before - pg_after))
            
            if inconsistent_tables:
                summary_lines.append("\n⚠️  INCONSISTENCIES DETECTED (for synced tables only):")
                for table_name, issue, actual, expected in inconsistent_tables:
                    if issue == 'count_mismatch':
                        summary_lines.append(f"  - {table_name}: PostgreSQL count ({actual}) does not match Access ({expected}) after sync")
                    else:
                        summary_lines.append(f"  - {table_name}: Deleted count ({actual}) does not match calculated deletion ({expected})")
            
        else:
            # Standard display (no sync_deleted performed)
            summary_lines.append("\nPer-table validation results:")
            summary_lines.append("-" * 70)
            summary_lines.append(f"{'Table Name':<30} {'Access':>10} {'PostgreSQL':>10} {'Status':>8}")
            summary_lines.append("-" * 70)
            
            for table_name, result in self.validation_results.items():
                source_count = result.get('source_count', 'N/A')
                target_count = fresh_counts.get(table_name, result.get('target_count', 'N/A'))
                matched = (source_count == target_count) if isinstance(source_count, int) and isinstance(target_count, int) else False
                status = "✓ PASS" if matched else "✗ FAIL"
                
                display_name = table_name[:29] if len(table_name) > 29 else table_name
                summary_lines.append(f"{display_name:<30} {source_count:>10} {target_count:>10} {status:>8}")
            
            summary_lines.append("-" * 70)
        
        total_tables = len(self.validation_results)
        passed_tables = 0
        for table_name, result in self.validation_results.items():
            source_count = result.get('source_count', 0)
            target_count = fresh_counts.get(table_name, result.get('target_count', 0))
            if source_count == target_count:
                passed_tables += 1
        failed_tables = total_tables - passed_tables
        total_deleted = sum(result.get('deleted', 0) for result in self.validation_results.values())
        
        summary_lines.append(f"\nTotal tables validated: {total_tables}")
        summary_lines.append(f"Passed: {passed_tables}")
        summary_lines.append(f"Failed: {failed_tables}")
        
        if total_deleted > 0:
            summary_lines.append(f"Total rows deleted from PostgreSQL: {total_deleted}")
        
        # Display empty tables information
        if self.empty_tables:
            summary_lines.append(f"\nEmpty tables (0 rows in source): {len(self.empty_tables)}")
            for table in self.empty_tables:
                summary_lines.append(f"  - {table}")
        
        summary_lines.append("=" * 110)
        
        # Print to console
        for line in summary_lines:
            print(line)
        
        # Write to log file
        for line in summary_lines:
            logger.info(line)
    
    # ========================================================================
    # SINGLE ROW INSERT (no batching)
    # ========================================================================
    
    def insert_row(self, safe_table_name, columns, row_data, conflict_clause=None):
        """Insert a single row into PostgreSQL.
        Returns: (inserted, result_type, details)
        result_type: 'SUCCESS', 'DUPLICATE_SKIPPED', 'FK_VIOLATION', 'ERROR'
        """
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - insert_row called with columns={columns}")
        
        if not row_data:
            return 0, 'ERROR', 'No data'
        
        # Build key string for logging (extract primary key values if possible)
        key_str = 'unknown'
        
        row_values = [convert_dao_value_to_postgresql_literal(val) for val in row_data]
        insert_sql = f"INSERT INTO {safe_table_name} ({', '.join(columns)}) VALUES ({', '.join(row_values)})"
        
        if conflict_clause:
            insert_sql += f" {conflict_clause}"
        
        try:
            self.pg_cursor.execute(insert_sql)
            self.pg_conn.commit()
            return 1, 'SUCCESS', key_str
        except Exception as e:
            self.pg_conn.rollback()
            if '23505' in str(e) or 'duplicate key' in str(e).lower():
                # PostgreSQL duplicate key error code 23505
                # Treat as success (row already in sync)
                if self.debug:
                    logger.debug(f"Duplicate key in INSERT for {safe_table_name}: {e}")
                return 1, 'DUPLICATE_SKIPPED', key_str
            elif 'foreign key constraint' in str(e).lower():
                logger.debug(f"Foreign key violation in row: {e}")
                return 0, 'FK_VIOLATION', key_str
            else:
                logger.warning(f"{insert_sql}")
                logger.warning(f"Failed to insert row: {e}")
                return 0, 'ERROR', key_str
    
    # ========================================================================
    # TABLE EXISTENCE CHECK
    # ========================================================================
    
    def table_exists_in_postgresql(self, safe_table_name):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - table_exists_in_postgresql called")
        
        try:
            unquoted_name = safe_table_name.strip('"')
            check_sql = """
                SELECT EXISTS (
                    SELECT 1 
                    FROM information_schema.tables 
                    WHERE table_schema = 'public' 
                    AND LOWER(table_name) = LOWER(%s)
                )
            """
            self.pg_cursor.execute(check_sql, (unquoted_name,))
            exists = self.pg_cursor.fetchone()[0]
            return exists
        except Exception as e:
            logger.error(f"Error checking if table exists: {e}")
            return False
    
    def row_exists_in_postgresql(self, safe_table_name, key_columns, key_values):
        """Check if a row exists in PostgreSQL by its key columns."""
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - row_exists_in_postgresql called")
        
        if not key_columns or not key_values:
            return False
        
        # If any key value is NULL, existence check is unreliable - return False
        if any(v is None for v in key_values):
            return False
        
        where_clause = ' AND '.join([f"{col} = %s" for col in key_columns])
        check_sql = f"SELECT COUNT(*) FROM {safe_table_name} WHERE {where_clause}"
        
        try:
            result = self.pg_sql_execute(check_sql, fetch_one=True, params=key_values)
            return result[0] > 0 if result else False
        except Exception as e:
            logger.warning(f"Error checking row existence: {e}")
            return False
    
    # ========================================================================
    # COLUMN TRANSFORMATION
    # ========================================================================
    
    def transform_column(self, table_name, column_name, value):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - Entering transform_column with table_name={repr(table_name)}, column_name={repr(column_name)}")
        
        transformations = self.parameters.get('transformations', [])
        
        if not isinstance(transformations, list):
            return value
        
        for trans in transformations:
            if not isinstance(trans, dict):
                continue
            
            trans_table = trans.get('table', '')
            if trans_table.lower() != table_name.lower():
                continue
            
            columns_list = trans.get('columns', [])
            if not isinstance(columns_list, list):
                continue
            
            for col_entry in columns_list:
                if not isinstance(col_entry, dict):
                    continue
                col_name = col_entry.get('name', '')
                pg_col_name = self.sanitise_token_for_postgresql(col_name)
                if pg_col_name.lower() != column_name.lower():
                    continue
                
                transform_type = col_entry.get('transform', '')
                
                if transform_type.lower() == 'drop':
                    return None
                elif transform_type.upper() == 'MMH3':
                    if value is not None:
                        return mmh3.hash(str(value))
                    return mmh3.hash('')
                elif transform_type.lower() == 'yearonly':
                    if isinstance(value, (datetime, date)):
                        return date(value.year, 1, 1)
                    return value
        
        return value
    
    # ========================================================================
    # COPY TABLE WITH ROW-BY-ROW PROCESSING (NO BATCHING)
    # ========================================================================
    
    def copy_table(self, table_info):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - copy_table called for {table_info['name']}")
        
        table_name = table_info['name']
        safe_table_name = table_info['safe_name']
        
        # Get accurate source row count from Access using SQL COUNT
        source_row_count = self.get_access_row_count(table_name)
        
        # Get the table definition for opening recordset
        tdef = self.dao_conn.TableDefs[table_name]
        
        if source_row_count == 0:
            msg = f"Table {table_name}: 0 rows (empty)"
            logger.info(msg)
            print(msg)
            self.empty_tables.append(table_name)
            self.validation_results[table_name] = {
                'source_count': 0,
                'target_count': 0,
                'matched': True,
                'difference': 0
            }
            return
        
        # Get target row count from PostgreSQL
        count_sql = f"SELECT COUNT(*) FROM {safe_table_name}"
        result = self.pg_sql_execute(count_sql, fetch_one=True)
        target_row_count = result[0] if result else 0
        
        # Check for non-volatile table optimization
        nonvolatile_tables = self.parameters.get('nonvolatile')
        if nonvolatile_tables is None:
            nonvolatile_tables = []
        is_nonvolatile = table_name in nonvolatile_tables
        
        # Handle non-volatile optimization with SLOW mode override
        if self.nonvolatile and is_nonvolatile:
            if source_row_count == target_row_count:
                if self.slow:
                    # SLOW mode overrides optimization - copy anyway
                    msg = f"Table {table_name} (non-volatile): row counts match (Access: {source_row_count}, PostgreSQL: {target_row_count}) - SLOW MODE: copying anyway (optimization disabled for debugging)"
                    logger.info(msg)
                    print(msg)
                    # Continue to normal copy below (do NOT return)
                else:
                    # Normal optimization - skip copy
                    msg = f"Table {table_name} (non-volatile): row counts match (Access: {source_row_count}, PostgreSQL: {target_row_count}) - skipping copy (--nonvolatile enabled)"
                    logger.info(msg)
                    print(msg)
                    # Store validation result without copying
                    self.validation_results[table_name] = {
                        'source_count': source_row_count,
                        'target_count': target_row_count,
                        'matched': True,
                        'difference': 0,
                        'nonvolatile_skipped': True
                    }
                    return
            else:
                msg = f"Table {table_name} (non-volatile): row counts differ (Access: {source_row_count}, PostgreSQL: {target_row_count}) - copying despite --nonvolatile"
                logger.info(msg)
                if self.verbose:
                    print(msg)
        
        has_unique_constraint = bool(table_info['primary_key'] or table_info['unique_keys'])
        
        start_position = 0
        rows_to_copy = 0
        
        if not has_unique_constraint:
            # Append-only table
            if target_row_count == 0:
                rows_to_copy = source_row_count
                start_position = 0
                msg = f"Table {table_name} (append-only): empty - copying all {rows_to_copy} rows"
                logger.info(msg)
                print(msg)
            elif target_row_count < source_row_count:
                rows_to_copy = source_row_count - target_row_count
                start_position = target_row_count
                msg = f"Table {table_name} (append-only): {target_row_count} existing rows - appending {rows_to_copy} new rows"
                logger.info(msg)
                print(msg)
            else:
                msg = f"Table {table_name} (append-only): up to date (Access: {source_row_count}, PostgreSQL: {target_row_count})"
                logger.info(msg)
                print(msg)
                self.validation_results[table_name] = {
                    'source_count': source_row_count,
                    'target_count': target_row_count,
                    'matched': source_row_count == target_row_count,
                    'difference': source_row_count - target_row_count
                }
                return
        else:
            # Has PK or UNIQUE - ALWAYS process ALL source rows
            rows_to_copy = source_row_count
            start_position = 0
            if target_row_count == 0:
                msg = f"Table {table_name}: empty - copying ALL {rows_to_copy} rows with UPSERT"
                if self.slow:
                    msg += " (SLOW MODE enabled - full sync)"
                logger.info(msg)
                print(msg)
            else:
                msg = f"Table {table_name}: {target_row_count} existing rows - syncing ALL {source_row_count} source rows with UPSERT"
                if self.slow:
                    msg += " (SLOW MODE enabled - full sync)"
                logger.info(msg)
                print(msg)
        
        violation_count = 0
        foreign_key_violations = 0
        duplicate_skipped_count = 0      # duplicate error but row exists (safe)
        duplicate_rejected_count = 0     # duplicate error and row does NOT exist (problem)
        rows_inserted = 0
        
        conflict_clause = None
        if has_unique_constraint:
            if table_info['primary_key']:
                conflict_columns = table_info['primary_key']
            elif table_info['unique_keys']:
                conflict_columns = table_info['unique_keys'][0]
            else:
                conflict_columns = None
            
            if conflict_columns:
                update_set = ', '.join([f"{col} = EXCLUDED.{col}" for col in [col['name'] for col in table_info['columns']]])
                conflict_clause = f"ON CONFLICT ({', '.join(conflict_columns)}) DO UPDATE SET {update_set}"
        
        # Open recordset for reading data
        recordset = tdef.OpenRecordset()
        
        if start_position > 0:
            try:
                recordset.Move(start_position)
            except Exception as e:
                logger.warning(f"Could not move to position {start_position}: {e}")
                start_position = 0
                rows_to_copy = source_row_count
        
        rows_processed = 0
        self.last_progress_time = time.time()
        copy_start_time = time.time()  # Record start time for ETA calculation
        
        while not recordset.EOF:
            row_data = []
            valid_columns = []
            
            for col in table_info['columns']:
                dao_val = recordset.Fields[col['original_name']].value
                py_val = convert_dao_value_to_python(dao_val)
                
                # Apply transformation
                transformed_val = self.transform_column(table_name, col['name'], py_val)
                
                # If transformation returns None, skip this column
                if transformed_val is None:
                    continue
                
                row_data.append(transformed_val)
                valid_columns.append(col['name'])
            
            rows_processed += 1
            stage = "Append" if not has_unique_constraint else "Sync"
            self.show_progress(table_name, rows_processed, rows_to_copy, stage, copy_start_time)
            
            # Insert the row if it has any columns
            if valid_columns:
                if conflict_clause and has_unique_constraint:
                    # Use ON CONFLICT for PostgreSQL UPSERT
                    inserted, result_type, details = self.insert_row(
                        safe_table_name, valid_columns, row_data, conflict_clause
                    )
                    if result_type == 'DUPLICATE_SKIPPED':
                        duplicate_skipped_count += 1
                        rows_inserted += inserted
                    elif result_type == 'DUPLICATE_REJECTED':
                        duplicate_rejected_count += 1
                        violation_count += 1
                    elif result_type == 'FK_VIOLATION':
                        foreign_key_violations += 1
                        violation_count += 1
                    else:
                        rows_inserted += inserted
                        violation_count += (1 - inserted)
                else:
                    inserted, result_type, details = self.insert_row(safe_table_name, valid_columns, row_data, None)
                    rows_inserted += inserted
                    violation_count += (1 - inserted)
            
            recordset.MoveNext()
        
        recordset.Close()
        
        self.show_progress(table_name, rows_to_copy, rows_to_copy, stage, copy_start_time)
        
        # Calculate elapsed time
        elapsed_time = time.time() - copy_start_time
        elapsed_str = self.format_elapsed(elapsed_time)
        
        if rows_inserted > 0:
            if not has_unique_constraint:
                msg = f"Completed {table_name}: Appended {rows_inserted} rows (elapsed: {elapsed_str})"
            else:
                msg = f"Completed {table_name}: Synced {rows_inserted} rows (elapsed: {elapsed_str})"
                if duplicate_skipped_count > 0:
                    msg += f" (duplicates skipped (exists): {duplicate_skipped_count})"
                if duplicate_rejected_count > 0:
                    msg += f" (duplicates rejected (missing): {duplicate_rejected_count})"
                if foreign_key_violations > 0:
                    msg += f" (FK violations: {foreign_key_violations})"
            logger.info(msg)
            print(msg)
        
        if rows_inserted > 0 or duplicate_skipped_count > 0 or duplicate_rejected_count > 0:
            self.validate_table_data(table_name, safe_table_name)
    
    # ========================================================================
    # SYSTEM TABLE FILTERING
    # ========================================================================
    
    def is_system_table(self, tdef):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - is_system_table called")
        
        table_name = decode_sketchy_utf16(tdef.Name)
        
        if table_name.lower().startswith('msys'):
            if self.verbose:
                logger.info(f"Filtering out MSys system table: {table_name}")
            return True
        
        if table_name.startswith('~'):
            if self.verbose:
                logger.info(f"Filtering out temporary table: {table_name}")
            return True
        
        DB_SYSTEM_OBJECT = -2147483648
        DB_HIDDEN_OBJECT = 1
        DB_ATTACHED_EXECUTE = 131072
        
        attributes = tdef.Attributes
        
        if attributes & DB_SYSTEM_OBJECT:
            if self.verbose:
                logger.info(f"Filtering out system object table: {table_name}")
            return True
        if attributes & DB_HIDDEN_OBJECT:
            if self.verbose:
                logger.info(f"Filtering out hidden table: {table_name}")
            return True
        if attributes & DB_ATTACHED_EXECUTE:
            if self.verbose:
                logger.info(f"Filtering out attached execute table: {table_name}")
            return True
        
        return False
    
    # ========================================================================
    # EXCLUSION CHECK
    # ========================================================================
    
    def is_table_excluded(self, table_name):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - is_table_excluded called with params: table_name={table_name}")
        
        excluded = self.parameters.get('excluded', [])
        if table_name in excluded:
            if self.verbose:
                logger.info(f"Table {table_name} is in excluded list - will be skipped")
            return True
        return False
    
    # ========================================================================
    # SANITIZATION FUNCTIONS
    # ========================================================================
    
    def sanitise_token_for_msaccess(self, token):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - sanitise_token_for_msaccess called with params: token={token}")
        
        if token is None:
            return ''
        token_str = str(token)
        if token_str.startswith('[') and token_str.endswith(']'):
            return token_str
        if token_str and token_str[0].isdigit():
            return f'[{token_str}]'
        if ' ' in token_str:
            return f'[{token_str}]'
        if re.search(r'[%\+\-/\.#!@$&*\(\)\.,;=<>?\\|`~^]', token_str):
            return f'[{token_str}]'
        if token_str.upper() in ms_access_reserved:
            return f'[{token_str}]'
        return token_str
    
    def sanitise_token_for_postgresql(self, token):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - sanitise_token_for_postgresql called with params: token={token}")
        
        if token is None or token == '':
            return ''
        
        # If already quoted, return as-is
        if token.startswith('"') and token.endswith('"'):
            return token
        
        # Remove Access brackets if present
        if token.startswith('[') and token.endswith(']'):
            token = token[1:-1]
        
        # Check if quoting is needed
        needs_quoting = False
        
        # Check for uppercase letters (preserve case sensitivity)
        if any(c.isupper() for c in token):
            needs_quoting = True
        
        # Check for special characters
        if '-' in token or ' ' in token:
            needs_quoting = True
        
        # Check for non-standard characters
        if re.search(r'[^a-zA-Z0-9_]', token):
            needs_quoting = True
        
        # Check if starts with digit
        if token and token[0].isdigit():
            needs_quoting = True
        
        # Check for reserved words
        if token.upper() in ['SELECT', 'INSERT', 'UPDATE', 'DELETE', 'FROM', 'WHERE', 'TABLE', 'COLUMN', 'INDEX', 'CONSTRAINT', 'PRIMARY', 'FOREIGN', 'KEY', 'UNIQUE', 'REFERENCES', 'CHECK', 'DEFAULT', 'NULL', 'NOT', 'AND', 'OR', 'IN', 'LIKE', 'BETWEEN', 'IS', 'AS', 'ON', 'JOIN', 'INNER', 'LEFT', 'RIGHT', 'OUTER', 'CROSS', 'UNION', 'ALL', 'DISTINCT', 'ORDER', 'GROUP', 'BY', 'HAVING', 'LIMIT', 'OFFSET', 'VALUES', 'SET', 'CREATE', 'ALTER', 'DROP', 'TRUNCATE', 'GRANT', 'REVOKE']:
            needs_quoting = True
        
        if needs_quoting:
            # Escape double quotes inside the token
            escaped = token.replace('"', '""')
            # Return quoted token preserving original case
            return f'"{escaped}"'
        
        # For tokens that don't need quoting, keep original case
        # PostgreSQL will fold to lowercase for unquoted identifiers
        return token
    
    def sanitise_keyname_for_postgresql(self, keyname):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - sanitise_keyname_for_postgresql called with params: keyname={keyname}")
        
        if not keyname:
            return ''
        keyname = str(keyname).lower()
        keyname = re.sub(r'[^a-z0-9_]', '_', keyname)
        if keyname and keyname[0].isdigit():
            keyname = '_' + keyname
        return keyname
    
    def normalise_index_name(self, table_name, index_name):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - normalise_index_name called with params: table_name={table_name}, index_name={index_name}")
        
        if not index_name:
            return ''
        table_name_clean = str(table_name).lower()
        index_name_clean = str(index_name).lower()
        prefix = f"{table_name_clean}_"
        if index_name_clean.startswith(prefix):
            return self.sanitise_keyname_for_postgresql(index_name)
        combined = f"{table_name_clean}_{index_name_clean}"
        normalised = self.sanitise_keyname_for_postgresql(combined)
        MAX_IDENTIFIER_LENGTH = 63
        if len(normalised.encode('utf-8')) > MAX_IDENTIFIER_LENGTH:
            truncated = normalised
            while len(truncated.encode('utf-8')) > MAX_IDENTIFIER_LENGTH and len(truncated) > 0:
                truncated = truncated[:-1]
            if truncated and truncated[-1] == '_':
                truncated = truncated[:-1]
            if not truncated:
                truncated = f"idx_{hashlib.md5(combined.encode()).hexdigest()[:16]}"
            normalised = truncated
        return normalised
    
    # ========================================================================
    # COLUMN NORMALIZATION
    # ========================================================================
    
    def normalize_column_name(self, column_name):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - normalize_column_name called with params: column_name={column_name}")
        
        if not column_name:
            return ""
        normalized = str(column_name).lower().strip()
        if normalized.startswith('['):
            normalized = normalized[1:]
        if normalized.endswith(']'):
            normalized = normalized[:-1]
        if normalized.startswith('"'):
            normalized = normalized[1:]
        if normalized.endswith('"'):
            normalized = normalized[:-1]
        normalized = re.sub(r'\s+', ' ', normalized)
        return normalized
    
    # ========================================================================
    # DATA TYPE CONVERSION
    # ========================================================================
    
    def convert_ms_datatype_postgresql(self, dao_type, field, column_name):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - convert_ms_datatype_postgresql called with params: dao_type={dao_type}, column_name={column_name}")
        
        type_map = {
                    1: 'BOOLEAN',
                    2: 'SMALLINT',
                    3: 'INTEGER',
                    4: 'BIGINT',
                    5: 'DECIMAL(19,4)',
                    6: 'REAL',
                    7: 'DOUBLE PRECISION',
                    8: 'TIMESTAMP',
                    9: 'BYTEA',
                    10: 'TEXT',
                    11: 'BYTEA',
                    12: 'TEXT',
                    15: 'UUID',
                    16: 'BIGINT',
                    20: 'DECIMAL',
                    21: 'BIGINT',
                    22: 'TIME',
                    23: 'TIMESTAMP',
                }
        pg_datatype = type_map.get(dao_type, 'TEXT')
        
    #    logger.info(f"convert_ms_datatype_postgresql: {column_name} [{dao_type}] --> {pg_datatype}")
        
        return pg_datatype
    
    # ========================================================================
    # CONNECTION MANAGEMENT
    # ========================================================================
    
    def open_DAO_connection(self):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - open_DAO_connection called")
        
        try:
            dao_engine = win32com.client.Dispatch("DAO.DBEngine.120")
            db_path = self.parameters.get('DAO', {}).get('database', 'Patients.mdb')
            full_path = os.path.abspath(db_path)
            self.dao_conn = dao_engine.OpenDatabase(full_path)
            logger.info(f"Connected to MS Access database: {full_path}")
            if self.verbose:
                print(f"Connected to MS Access database: {full_path}")
            return True
        except Exception as e:
            logger.error(f"Failed to connect to MS Access database: {e}")
            raise
    
    def open_postgresql_connection(self):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - open_postgresql_connection called")
        
        pg_params = self.parameters.get('postgresql', {})
        required = ['host', 'port', 'database', 'user', 'password']
        for param in required:
            if param not in pg_params or pg_params[param] is None or pg_params[param] == '':
                logger.error(f"Missing PostgreSQL parameter: {param}")
                return False
        try:
            self.pg_conn = psycopg2.connect(
                host=pg_params['host'],
                port=int(pg_params['port']),
                database=pg_params['database'],
                user=pg_params['user'],
                password=pg_params['password']
            )
            self.pg_cursor = self.pg_conn.cursor()
            logger.info(f"Connected to PostgreSQL database: {pg_params['database']}")
            if self.verbose:
                print(f"Connected to PostgreSQL database: {pg_params['database']}")
            return True
        except psycopg2.OperationalError as e:
            logger.error(f"PostgreSQL connection failed: {e}")
            logger.error("Check that:")
            logger.error("1. PostgreSQL server is running")
            logger.error("2. Host and port are correct")
            logger.error("3. Username and password are correct")
            logger.error("4. Database exists")
            raise
    
    def open_postgresql_connection_master(self):
        """Connect to PostgreSQL using 'postgres' database for testing purposes only."""
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - open_postgresql_connection_master called")
        
        pg_params = self.parameters.get('postgresql', {})
        required = ['host', 'port', 'user', 'password']
        for param in required:
            if param not in pg_params or pg_params[param] is None or pg_params[param] == '':
                logger.error(f"Missing PostgreSQL parameter: {param}")
                return False
        
        try:
            self.pg_conn = psycopg2.connect(
                host=pg_params['host'],
                port=int(pg_params['port']),
                database='postgres',
                user=pg_params['user'],
                password=pg_params['password']
            )
            self.pg_cursor = self.pg_conn.cursor()
            logger.info("Connected to PostgreSQL postgres database for testing")
            if self.verbose:
                print("Connected to PostgreSQL postgres database for testing")
            return True
        except psycopg2.OperationalError as e:
            logger.error(f"PostgreSQL connection failed: {e}")
            raise
    
    def open_connections(self):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - open_connections called")
        
        self.open_DAO_connection()
        self.open_postgresql_connection()
    
    def close_connections(self):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - close_connections called")
        
        if self.pg_cursor:
            self.pg_cursor.close()
        if self.pg_conn:
            self.pg_conn.close()
        if self.dao_conn:
            self.dao_conn.Close()
        logger.info("Connections closed")
    
    # ========================================================================
    # SQL EXECUTION
    # ========================================================================
    
    def pg_sql_execute(self, sql, fetch_one=False, fetch_all=False, params=None):
        if self.trace:
            frame = inspect.currentframe()
            sql_preview = sql[:500] + "..." if len(sql) > 500 else sql
            logger.info(f"Line {frame.f_lineno} - pg_sql_execute called with sql length: {len(sql)} chars, fetch_one={fetch_one}, fetch_all={fetch_all}")
            if self.debug:
                logger.debug(f"SQL: {sql_preview}")
        
        if self.debug:
            safe_sql = re.sub(r"PASSWORD\s*=\s*'[^']*'", "PASSWORD='***'", sql, flags=re.IGNORECASE)
            logger.debug(f"Executing SQL: {safe_sql[:500]}")
        
        try:
            if params:
                self.pg_cursor.execute(sql, params)
            else:
                self.pg_cursor.execute(sql)
            if fetch_one:
                return self.pg_cursor.fetchone()
            if fetch_all:
                return self.pg_cursor.fetchall()
            return None
        except Exception as e:
            logger.error(f"SQL execution failed: {e}")
            logger.error(f"SQL statement (first 500 chars): {sql[:500]}")
            print(f"SQL execution failed: {e}")
            if self.debug:
                print(f"Failed SQL: {sql[:500]}")
            raise
    
    # ========================================================================
    # INDEX LOADING
    # ========================================================================
    
    def load_indexes(self, table_name):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - load_indexes called for {table_name}")
        
        indexes = []
        tdef = self.dao_conn.TableDefs[table_name]
        for i in range(tdef.Indexes.Count):
            idx = tdef.Indexes[i]
            index_info = {
                'name': decode_sketchy_utf16(idx.Name),
                'columns': [],
                'primary': False,
                'unique': False
            }
            has_attributes = hasattr(idx, 'Attributes')
            has_primary = hasattr(idx, 'Primary')
            has_unique = hasattr(idx, 'Unique')
            
            if not has_attributes and not has_primary and not has_unique:
                logger.error(f"FATAL: Index '{idx.Name}' has no properties")
                sys.exit(1)
            
            if has_attributes:
                attrs = idx.Attributes
                if attrs & 1:
                    index_info['primary'] = True
                if attrs & 2:
                    index_info['unique'] = True
            
            if has_primary:
                try:
                    if idx.Primary:
                        index_info['primary'] = True
                except:
                    pass
            
            if has_unique:
                try:
                    if idx.Unique:
                        index_info['unique'] = True
                except:
                    pass
            
            for j in range(idx.Fields.Count):
                fld = idx.Fields[j]
                index_info['columns'].append(decode_sketchy_utf16(fld.Name))
            
            indexes.append(index_info)
        
        return indexes
    
    # ========================================================================
    # TABLE LOADING
    # ========================================================================
    
    def load_table(self, table_name):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - load_table called for {table_name}")
        
        tdef = self.dao_conn.TableDefs[table_name]
        
        table_info = {
            'name': table_name,
            'safe_name': self.sanitise_token_for_postgresql(table_name),
            'columns': [],
            'indexes': self.load_indexes(table_name),
            'primary_key': [],
            'unique_keys': []
        }
        
        col_mapping = self.parameters.get('column_mapping', {}).get(table_name, {})
        
        for i in range(tdef.Fields.Count):
            fld = tdef.Fields[i]
            orig_name = decode_sketchy_utf16(fld.Name)
            normalized = self.normalize_column_name(orig_name)
            safe_name = col_mapping.get(normalized, self.sanitise_token_for_postgresql(orig_name))
            
            col_info = {
                'original_name': orig_name,
                'name': safe_name,
                'type': self.convert_ms_datatype_postgresql(fld.Type, fld, orig_name),
                'required': (fld.Required if hasattr(fld, 'Required') else False)
            }
            table_info['columns'].append(col_info)
        
        for idx in table_info.get('indexes', []):
            if idx.get('primary'):
                for col_name in idx['columns']:
                    for col in table_info['columns']:
                        if col['original_name'] == col_name or \
                           self.normalize_column_name(col['original_name']) == self.normalize_column_name(col_name):
                            table_info['primary_key'].append(col['name'])
                            break
                break
        
        for idx in table_info.get('indexes', []):
            if idx.get('unique') and not idx.get('primary'):
                unique_columns = []
                for col_name in idx['columns']:
                    for col in table_info['columns']:
                        if col['original_name'] == col_name or \
                           self.normalize_column_name(col['original_name']) == self.normalize_column_name(col_name):
                            unique_columns.append(col['name'])
                            break
                if unique_columns:
                    table_info['unique_keys'].append(unique_columns)
        
        return table_info
    
    # ========================================================================
    # TABLE DISCOVERY
    # ========================================================================
    
    def get_all_tables_to_process(self):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - get_all_tables_to_process called")
        
        tables_config = self.parameters.get('tables')
        if tables_config is None:
            tables_config = []
        
        if tables_config and len(tables_config) > 0:
            logger.info("Using 'tables' section from configuration")
            table_names = []
            for entry in tables_config:
                if isinstance(entry, dict):
                    table_names.append(entry.get('name', ''))
                else:
                    table_names.append(str(entry))
            table_names = [t for t in table_names if t]
        else:
            logger.info("No 'tables' section in configuration - auto-discovering all non-system tables")
            table_names = []
            for i in range(self.dao_conn.TableDefs.Count):
                tdef = self.dao_conn.TableDefs[i]
                if not self.is_system_table(tdef):
                    table_names.append(decode_sketchy_utf16(tdef.Name))
            logger.info(f"Auto-discovered {len(table_names)} non-system tables")
            if self.debug:
                logger.debug(f"Auto-discovered tables before exclusion: {table_names}")
        
        result = [t for t in table_names if not self.is_table_excluded(t)]
        
        if self.verbose:
            logger.info(f"After exclusion filtering: {len(result)} tables to process: {result}")
        
        if not result:
            logger.warning("No tables to process after filtering")
        
        return result
    
    # ========================================================================
    # PRIMARY KEY AND UNIQUE CONSTRAINT CHECKS
    # ========================================================================
    
    def check_primary_key(self, table_name, columns):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - check_primary_key called with params: table_name={table_name}, columns={columns}")
        
        safe_table_name = self.sanitise_token_for_postgresql(table_name)
        unquoted_table_name = safe_table_name.strip('"')
        columns_lower = [col.lower() for col in columns]
        
        pk_columns_sql = f"""
            SELECT LOWER(a.attname) FROM pg_index i
            JOIN pg_attribute a ON a.attrelid = i.indrelid AND a.attnum = ANY(i.indkey)
            JOIN pg_class c ON c.oid = i.indrelid
            JOIN pg_namespace n ON n.oid = c.relnamespace
            WHERE c.relname = '{escape_postgresql_string(unquoted_table_name)}'
            AND n.nspname = 'public'
            AND i.indisprimary
            ORDER BY a.attnum
        """
        pk_columns = self.pg_sql_execute(pk_columns_sql, fetch_all=True)
        pk_column_names = [row[0] for row in pk_columns] if pk_columns else []
        return set(columns_lower) == set(pk_column_names)
    
    def check_unique_constraint_only(self, table_name, columns):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - check_unique_constraint_only called with params: table_name={table_name}, columns={columns}")
        
        if self.check_primary_key(table_name, columns):
            return False
        safe_table_name = self.sanitise_token_for_postgresql(table_name)
        unquoted_table_name = safe_table_name.strip('"')
        columns_lower = [col.lower() for col in columns]
        
        unique_check_sql = f"""
            SELECT indexname, indexdef 
            FROM pg_indexes 
            WHERE LOWER(tablename) = LOWER('{escape_postgresql_string(unquoted_table_name)}')
            AND schemaname = 'public'
            AND indexdef LIKE 'CREATE UNIQUE INDEX%'
        """
        unique_indexes = self.pg_sql_execute(unique_check_sql, fetch_all=True)
        for idx_name, idx_def in unique_indexes or []:
            match = re.search(r'\(([^)]+)\)', idx_def)
            if match:
                index_columns = []
                for col_part in match.group(1).split(','):
                    col_part = col_part.strip()
                    if col_part.startswith('"') and col_part.endswith('"'):
                        col_part = col_part[1:-1]
                    index_columns.append(col_part.lower())
                if set(columns_lower) == set(index_columns):
                    return True
        return False
    
    def determine_correct_direction(self, table_a, columns_a, table_b, columns_b):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - determine_correct_direction called with params: table_a={table_a}, columns_a={columns_a}, table_b={table_b}, columns_b={columns_b}")
        
        a_is_pk = self.check_primary_key(table_a, columns_a)
        b_is_pk = self.check_primary_key(table_b, columns_b)
        a_is_unique = self.check_unique_constraint_only(table_a, columns_a)
        b_is_unique = self.check_unique_constraint_only(table_b, columns_b)
        
        logger.info(f"Direction analysis for {table_a}({columns_a}) <-> {table_b}({columns_b})")
        logger.info(f"  {table_a}: PRIMARY KEY={a_is_pk}, UNIQUE={a_is_unique}")
        logger.info(f"  {table_b}: PRIMARY KEY={b_is_pk}, UNIQUE={b_is_unique}")
        
        if a_is_pk and not b_is_pk:
            logger.info(f"  Decision: {table_a} has PRIMARY KEY -> Parent, {table_b} -> Child")
            return (table_b, columns_b, table_a, columns_a)
        elif b_is_pk and not a_is_pk:
            logger.info(f"  Decision: {table_b} has PRIMARY KEY -> Parent, {table_a} -> Child")
            return (table_a, columns_a, table_b, columns_b)
        
        if a_is_unique and not b_is_unique:
            logger.info(f"  Decision: {table_a} has UNIQUE -> Parent, {table_b} -> Child")
            return (table_b, columns_b, table_a, columns_a)
        elif b_is_unique and not a_is_unique:
            logger.info(f"  Decision: {table_b} has UNIQUE -> Parent, {table_a} -> Child")
            return (table_a, columns_a, table_b, columns_b)
        
        if a_is_pk or b_is_pk or a_is_unique or b_is_unique:
            logger.warning(f"  Decision: Both tables have uniqueness - direction ambiguous")
            logger.warning(f"  Preserving original: {table_a} -> {table_b}")
            return (table_a, columns_a, table_b, columns_b)
        else:
            logger.warning(f"  Decision: Neither table has PRIMARY KEY or UNIQUE constraint")
            return None
    
    def is_direction_valid(self, base_table, base_columns, reference_table, reference_columns):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - is_direction_valid called with params: base_table={base_table}, base_columns={base_columns}, reference_table={reference_table}, reference_columns={reference_columns}")
        
        correct = self.determine_correct_direction(base_table, base_columns, reference_table, reference_columns)
        if correct is None:
            return None
        correct_child, correct_child_cols, correct_parent, correct_parent_cols = correct
        if base_table == correct_child and reference_table == correct_parent:
            logger.info(f"  Given direction {base_table} -> {reference_table} is VALID")
            return True
        elif base_table == correct_parent and reference_table == correct_child:
            logger.info(f"  Given direction {base_table} -> {reference_table} is BACKWARDS")
            return False
        return None
    
    def attempt_direction_correction(self, fk_name, fk_info, is_from_config):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - attempt_direction_correction called with params: fk_name={fk_name}, is_from_config={is_from_config}")
        
        base_table = fk_info['base_table']
        reference_table = fk_info['reference_table']
        base_columns = fk_info['base_columns']
        reference_columns = fk_info['reference_columns']
        
        logger.info(f"Processing foreign key: {fk_name}")
        correct = self.determine_correct_direction(base_table, base_columns, reference_table, reference_columns)
        if correct is None:
            return None
        correct_child, correct_child_cols, correct_parent, correct_parent_cols = correct
        
        if base_table == correct_child and reference_table == correct_parent:
            logger.info(f"  Foreign key {fk_name} direction is correct")
            return None
        elif base_table == correct_parent and reference_table == correct_child:
            logger.warning(f"  Foreign key {fk_name} direction is INCORRECT - needs swapping")
            logger.info(f"  Original: {base_table}({base_columns}) -> {reference_table}({reference_columns})")
            logger.info(f"  Corrected: {correct_child}({correct_child_cols}) -> {correct_parent}({correct_parent_cols})")
            return {
                'base_table': correct_child,
                'reference_table': correct_parent,
                'base_columns': correct_child_cols,
                'reference_columns': correct_parent_cols
            }
        return None
    
    # ========================================================================
    # CHECK REFERENCE TABLE UNIQUENESS (kept for possible future use)
    # ========================================================================
    
    def check_reference_table_has_uniqueness(self, reference_table, reference_columns):
        """Check if reference table has a PRIMARY KEY or UNIQUE constraint on the specified columns.
        Returns: tuple (has_constraint, constraint_type, constraint_name)
        """
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - check_reference_table_has_uniqueness called for {reference_table}.{reference_columns}")
        
        safe_ref_table = self.sanitise_token_for_postgresql(reference_table)
        unquoted_table = safe_ref_table.strip('"')
        columns_lower = [col.lower().strip('"') for col in reference_columns]
        
        # Check for PRIMARY KEY
        pk_sql = f"""
            SELECT LOWER(a.attname), c.conname
            FROM pg_constraint c
            JOIN pg_class t ON t.oid = c.conrelid
            JOIN pg_namespace n ON n.oid = t.relnamespace
            JOIN pg_attribute a ON a.attrelid = t.oid AND a.attnum = ANY(c.conkey)
            WHERE t.relname = '{escape_postgresql_string(unquoted_table)}'
            AND n.nspname = 'public'
            AND c.contype = 'p'
            ORDER BY a.attnum
        """
        pk_columns = self.pg_sql_execute(pk_sql, fetch_all=True)
        if pk_columns:
            pk_col_names = [row[0] for row in pk_columns]
            pk_constraint_name = pk_columns[0][1] if pk_columns else None
            if set(columns_lower) == set(pk_col_names):
                logger.info(f"Reference table {reference_table} has PRIMARY KEY on {reference_columns}")
                return (True, 'PRIMARY KEY', pk_constraint_name)
        
        # Check for UNIQUE constraint
        unique_sql = f"""
            SELECT LOWER(a.attname), c.conname
            FROM pg_constraint c
            JOIN pg_class t ON t.oid = c.conrelid
            JOIN pg_namespace n ON n.oid = t.relnamespace
            JOIN pg_attribute a ON a.attrelid = t.oid AND a.attnum = ANY(c.conkey)
            WHERE t.relname = '{escape_postgresql_string(unquoted_table)}'
            AND n.nspname = 'public'
            AND c.contype = 'u'
        """
        unique_constraints = self.pg_sql_execute(unique_sql, fetch_all=True)
        
        # Group by constraint name
        constraints = {}
        for col_name, con_name in unique_constraints or []:
            if con_name not in constraints:
                constraints[con_name] = []
            constraints[con_name].append(col_name)
        
        for con_name, col_list in constraints.items():
            if set(columns_lower) == set(col_list):
                logger.info(f"Reference table {reference_table} has UNIQUE constraint {con_name} on {reference_columns}")
                return (True, 'UNIQUE', con_name)
        
        # Check for UNIQUE INDEX (not a constraint)
        unique_index_sql = f"""
            SELECT indexname, indexdef
            FROM pg_indexes
            WHERE LOWER(tablename) = LOWER('{escape_postgresql_string(unquoted_table)}')
            AND schemaname = 'public'
            AND indexdef LIKE 'CREATE UNIQUE INDEX%'
        """
        unique_indexes = self.pg_sql_execute(unique_index_sql, fetch_all=True)
        for idx_name, idx_def in unique_indexes or []:
            match = re.search(r'\(([^)]+)\)', idx_def)
            if match:
                index_columns = []
                for col_part in match.group(1).split(','):
                    col_part = col_part.strip()
                    if col_part.startswith('"') and col_part.endswith('"'):
                        col_part = col_part[1:-1]
                    index_columns.append(col_part.lower())
                if set(columns_lower) == set(index_columns):
                    logger.info(f"Reference table {reference_table} has UNIQUE INDEX {idx_name} on {reference_columns}")
                    return (True, 'UNIQUE INDEX', idx_name)
        
        logger.warning(f"Reference table {reference_table} has no PRIMARY KEY or UNIQUE constraint on columns {reference_columns}")
        return (False, None, None)
    
    # ========================================================================
    # CREATE UNIQUE CONSTRAINT ON BASE TABLE (PRIVATE HELPER)
    # ========================================================================
    # NOTE: PostgreSQL does NOT require the child side of a foreign key to be unique.
    # Only the parent side (referenced table) needs a PRIMARY KEY or UNIQUE constraint.
    # Therefore, this method should NOT be called. It is kept for completeness but
    # the ensure_uniqueness_on_base_table method has been modified to skip creation
    # of unnecessary unique constraints on child tables.
    # ========================================================================
    
    def _create_unique_constraint_on_base_table(self, base_table, base_columns, fk_name):
        """Create a UNIQUE constraint on the base table (child) for the columns used in the foreign key.
        NOTE: This method should not be called as PostgreSQL does not require uniqueness
        on the child side of a foreign key. It is kept for completeness but will log a warning.
        """
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - _create_unique_constraint_on_base_table called")
        
        logger.warning(f"NOT CREATING UNIQUE constraint on {base_table} - PostgreSQL does not require uniqueness on the child side of a foreign key. The foreign key will reference the parent's PRIMARY KEY directly.")
        return None
    
    # ========================================================================
    # CREATE PRIMARY KEY ON BASE TABLE (PRIVATE HELPER)
    # ========================================================================
    
    def _create_primary_key_on_base_table(self, base_table, base_columns, fk_name):
        """Create a PRIMARY KEY constraint on the base table (child) for the columns used in the foreign key.
        Note: This will fail if any of the columns contain NULL values or if a primary key already exists.
        """
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - _create_primary_key_on_base_table called")
        
        safe_base_table = self.sanitise_token_for_postgresql(base_table)
        safe_columns = [self.sanitise_token_for_postgresql(col) for col in base_columns]
        
        # Generate a constraint name based on the FK name
        constraint_name = self.sanitise_keyname_for_postgresql(f"pk_{fk_name}")
        
        # Check if a PRIMARY KEY constraint already exists with this name
        check_sql = f"""
            SELECT 1 FROM pg_constraint c
            JOIN pg_class t ON t.oid = c.conrelid
            JOIN pg_namespace n ON n.oid = t.relnamespace
            WHERE t.relname = '{escape_postgresql_string(safe_base_table.strip('"'))}'
            AND n.nspname = 'public'
            AND c.contype = 'p'
            AND c.conname = '{escape_postgresql_string(constraint_name)}'
        """
        exists = self.pg_sql_execute(check_sql, fetch_one=True)
        if exists:
            logger.info(f"PRIMARY KEY constraint {constraint_name} already exists on {base_table}")
            return constraint_name
        
        # Create the PRIMARY KEY constraint
        create_sql = f"""
            ALTER TABLE {safe_base_table}
            ADD CONSTRAINT {constraint_name}
            PRIMARY KEY ({', '.join(safe_columns)})
        """
        
        try:
            self.pg_sql_execute(create_sql)
            self.pg_conn.commit()
            logger.info(f"Created PRIMARY KEY constraint {constraint_name} on {base_table}({', '.join(base_columns)})")
            print(f"  Created PRIMARY KEY constraint on {base_table} for foreign key reference")
            return constraint_name
        except Exception as e:
            self.pg_conn.rollback()
            logger.error(f"Failed to create PRIMARY KEY constraint on {base_table}: {e}")
            raise
    
    # ========================================================================
    # ENSURE UNIQUENESS ON BASE TABLE (PUBLIC METHOD)
    # ========================================================================
    
    def ensure_uniqueness_on_base_table(self, base_table, base_columns, fk_name):
        """Ensure a uniqueness constraint (PRIMARY KEY if possible, otherwise UNIQUE) 
        on the base table columns required for the foreign key.
        
        NOTE: PostgreSQL does NOT require the child side of a foreign key to be unique.
        Only the parent side (referenced table) needs uniqueness. Therefore, this method
        now returns None (no action) and logs a warning instead of creating unnecessary
        unique constraints.
        """
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - ensure_uniqueness_on_base_table called")
        
        logger.warning(f"Foreign key {fk_name}: PostgreSQL does NOT require uniqueness on child table '{base_table}' columns {base_columns}. Skipping automatic creation of unique constraint.")
        return None
    
    # ========================================================================
    # ADJUST MS ACCESS SCHEMA
    # ========================================================================
    
    def adjust_ms_access_schema(self):
        """Adjust MS Access schema: add AutoNumber primary key to tables without a PK."""
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - adjust_ms_access_schema called")
        
        print("\n" + "=" * 60)
        print("ADJUST MS ACCESS SCHEMA")
        print("=" * 60)
        
        # Open only DAO connection (no PostgreSQL needed)
        try:
            self.open_DAO_connection()
            print("✓ Connected to MS Access database.")
            logger.info("adjust_ms_access_schema: DAO connection successful")
        except Exception as e:
            print(f"✗ Failed to connect to MS Access database: {e}")
            logger.error(f"adjust_ms_access_schema: DAO connection failed - {e}")
            sys.exit(1)
        
        # Get all user tables (non-system)
        table_names = []
        for i in range(self.dao_conn.TableDefs.Count):
            tdef = self.dao_conn.TableDefs[i]
            if not self.is_system_table(tdef):
                table_names.append(decode_sketchy_utf16(tdef.Name))
        
        print(f"\nFound {len(table_names)} user tables. Checking for primary keys...")
        logger.info(f"adjust_ms_access_schema: Found {len(table_names)} user tables")
        
        modified_count = 0
        skipped_count = 0
        error_count = 0
        
        for table_name in table_names:
            tdef = self.dao_conn.TableDefs[table_name]
            
            # Check if table already has a primary key
            has_primary_key = False
            for idx in tdef.Indexes:
                if idx.Primary:
                    has_primary_key = True
                    break
            
            if has_primary_key:
                print(f"  ⊘ {table_name}: already has PRIMARY KEY - skipping")
                logger.info(f"adjust_ms_access_schema: {table_name} already has PK, skipping")
                skipped_count += 1
                continue
            
            # Check if table already has an AutoNumber field
            existing_autonumber_column = None
            for fld in tdef.Fields:
                if fld.Attributes & 16:  # 16 = dbAutoIncrField
                    existing_autonumber_column = fld.Name
                    break
            
            if existing_autonumber_column:
                # Use existing AutoNumber column as primary key
                print(f"  🔑 {table_name}: using existing column '{existing_autonumber_column}' as PRIMARY KEY")
                logger.info(f"adjust_ms_access_schema: {table_name} using existing column {existing_autonumber_column}")
                
                try:
                    # Check if this column already has a primary key
                    has_pk_on_column = False
                    for idx in tdef.Indexes:
                        if idx.Primary and idx.Fields.Count == 1:
                            idx_field_name = idx.Fields(0).Name
                            if idx_field_name == existing_autonumber_column:
                                has_pk_on_column = True
                                break
                    
                    if not has_pk_on_column:
                        # Create primary key on the existing AutoNumber column
                        idx = tdef.CreateIndex("PrimaryKey")
                        idx.Primary = True
                        idx.Fields.Append(idx.CreateField(existing_autonumber_column))
                        tdef.Indexes.Append(idx)
                        print(f"    ✓ Successfully added primary key on '{existing_autonumber_column}'")
                        logger.info(f"adjust_ms_access_schema: Added PK on existing column in {table_name}")
                        modified_count += 1
                    else:
                        print(f"    ⊘ Column '{existing_autonumber_column}' already has PRIMARY KEY - skipping")
                        skipped_count += 1
                except Exception as e:
                    print(f"    ✗ ERROR on {table_name}: {e}")
                    logger.error(f"adjust_ms_access_schema: Failed to add PK on existing column in {table_name} - {e}")
                    error_count += 1
                continue  # Skip to next table
            
            # Generate unique column name based on table name
            col_name = f"{table_name}_replicator_id".lower()
            col_name = re.sub(r'[^a-z0-9_]', '_', col_name)
            
            if len(col_name) > 64:
                col_name = col_name[:64]
            
            # Check if column already exists
            column_exists = False
            try:
                _ = tdef.Fields(col_name)
                column_exists = True
            except Exception:
                column_exists = False
            
            if column_exists:
                print(f"  ⚠ {table_name}: column '{col_name}' already exists - skipping")
                logger.warning(f"adjust_ms_access_schema: {table_name} column {col_name} already exists")
                skipped_count += 1
                continue
            
            # Table lacks a primary key - add AutoNumber column and make it PK
            print(f"  ➕ {table_name}: adding AutoNumber primary key column '{col_name}'")
            logger.info(f"adjust_ms_access_schema: Adding PK to {table_name} with column {col_name}")
            
            try:
                # Step 1: Add AutoNumber column
                fld = tdef.CreateField(col_name, 4)  # 4 = dbLong (Long Integer)
                fld.Attributes = 16  # 16 = dbAutoIncrField (AutoNumber)
                tdef.Fields.Append(fld)
                
                # Step 2: Create primary key index on the new column
                idx = tdef.CreateIndex("PrimaryKey")
                idx.Primary = True
                idx.Fields.Append(idx.CreateField(col_name))
                tdef.Indexes.Append(idx)
                
                modified_count += 1
                print(f"    ✓ Successfully added primary key on '{col_name}'")
                logger.info(f"adjust_ms_access_schema: Successfully added PK to {table_name}")
            except Exception as e:
                print(f"    ✗ ERROR on {table_name}: {e}")
                logger.error(f"adjust_ms_access_schema: Failed to add PK to {table_name} - {e}")
                error_count += 1
        
        # Close the DAO connection
        self.close_connections()
        
        print("\n" + "=" * 60)
        print(f"SUMMARY:")
        print(f"  Tables modified:   {modified_count}")
        print(f"  Tables skipped:    {skipped_count}")
        print(f"  Errors:            {error_count}")
        print("=" * 60)
        logger.info(f"adjust_ms_access_schema completed: {modified_count} modified, {skipped_count} skipped, {error_count} errors")
    
    # ========================================================================
    # SYNC DELETED TABLES
    # ========================================================================
    
    def sync_deleted_tables(self, tables_to_process):
        """Synchronize deleted records by comparing row counts between Access and PostgreSQL.
        If counts differ, call _sync_deleted_table() for that table."""
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - sync_deleted_tables called with {len(tables_to_process)} tables")
        
        print("\n" + "=" * 60)
        print("SYNC DELETED RECORDS")
        print("=" * 60)
        
        if self.slow:
            print("⚠️  SLOW MODE enabled - will process all tables regardless of row counts.")
            logger.info("sync_deleted_tables: slow mode enabled")
        else:
            print("Fast mode - will only process tables where row counts differ.")
            logger.info("sync_deleted_tables: fast mode enabled")
        
        print()
        
        tables_processed = 0
        tables_skipped = 0
        
        for table_info in tables_to_process:
            table_name = table_info['name']
            safe_table_name = table_info['safe_name']
            
            # Get row count from MS Access
            try:
                access_count = self.get_access_row_count(table_name)
            except Exception as e:
                print(f"  ✗ {table_name}: Failed to get Access row count - {e}")
                logger.error(f"sync_deleted_tables: Failed to get Access count for {table_name} - {e}")
                continue
            
            # Get row count from PostgreSQL
            try:
                count_sql = f"SELECT COUNT(*) FROM {safe_table_name}"
                result = self.pg_sql_execute(count_sql, fetch_one=True)
                pg_count = result[0] if result else 0
            except Exception as e:
                print(f"  ✗ {table_name}: Failed to get PostgreSQL row count - {e}")
                logger.error(f"sync_deleted_tables: Failed to get PostgreSQL count for {table_name} - {e}")
                continue
            
            # Diagnostic logging to compare copy phase vs sync phase
            original_validation = self.validation_results.get(table_name, {})
            logger.info(f"sync_deleted_tables: {table_name} - Copy phase: Access={original_validation.get('source_count')}, PostgreSQL={original_validation.get('target_count')}")
            logger.info(f"sync_deleted_tables: {table_name} - Current check: Access={access_count}, PostgreSQL={pg_count}")
            
            print(f"  {table_name}: Access={access_count}, PostgreSQL={pg_count}")
            logger.info(f"sync_deleted_tables: {table_name} Access={access_count}, PostgreSQL={pg_count}")
            
            # Decide whether to process this table
            if self.slow:
                # Slow mode: process all tables regardless of counts
                print(f"    → SLOW MODE: processing {table_name}")
                self._sync_deleted_table(table_name, safe_table_name, access_count, pg_count)
                tables_processed += 1
            elif access_count != pg_count:
                # Fast mode: only process when counts differ
                print(f"    → Counts differ - processing {table_name}")
                self._sync_deleted_table(table_name, safe_table_name, access_count, pg_count)
                tables_processed += 1
            else:
                # Fast mode: counts match, skip
                print(f"    → Counts match - skipping (no deleted records to sync)")
                tables_skipped += 1
        
        print("\n" + "=" * 60)
        print(f"SYNC DELETED SUMMARY:")
        print(f"  Tables processed: {tables_processed}")
        print(f"  Tables skipped:   {tables_skipped}")
        print("=" * 60)
        logger.info(f"sync_deleted_tables completed: {tables_processed} processed, {tables_skipped} skipped")
    
    # ========================================================================
    # SYNC DELETED TABLE (PRIVATE METHOD - BATCH PROCESSING WITH KEYSET PAGINATION)
    # ========================================================================
    
    def _sync_deleted_table(self, table_name, safe_table_name, access_count, pg_count):
        """Delete rows from PostgreSQL that no longer exist in MS Access.
        Uses PRIMARY KEY or first UNIQUE constraint to identify matching rows.
        Processes in batches using keyset pagination to avoid memory issues.
        """
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - _sync_deleted_table called for {table_name}")
        
        print(f"      Syncing {table_name}: PostgreSQL has {pg_count} rows, Access has {access_count} rows")
        logger.info(f"_sync_deleted_table: Processing {table_name} (Access: {access_count}, PostgreSQL: {pg_count})")
        
        # Store original pg_count for validation summary
        original_pg_count = pg_count
        
        # Record start time for elapsed calculation
        sync_start_time = time.time()
        
        # Step 1: Identify the key columns to use for matching (PK or first UNIQUE)
        key_columns = []
        key_column_names = []
        
        # First, try to find PRIMARY KEY
        pk_sql = f"""
            SELECT a.attname
            FROM pg_index i
            JOIN pg_attribute a ON a.attrelid = i.indrelid AND a.attnum = ANY(i.indkey)
            JOIN pg_class c ON c.oid = i.indrelid
            JOIN pg_namespace n ON n.oid = c.relnamespace
            WHERE c.relname = '{escape_postgresql_string(safe_table_name.strip('"'))}'
            AND n.nspname = 'public'
            AND i.indisprimary
            ORDER BY a.attnum
        """
        pk_columns = self.pg_sql_execute(pk_sql, fetch_all=True)
        
        if pk_columns:
            key_columns = [row[0] for row in pk_columns]
            key_column_names = [self.sanitise_token_for_postgresql(col) for col in key_columns]
            logger.info(f"_sync_deleted_table: {table_name} using PRIMARY KEY: {', '.join(key_columns)}")
        else:
            # No PRIMARY KEY - look for first UNIQUE constraint
            unique_sql = f"""
                SELECT a.attname
                FROM pg_constraint c
                JOIN pg_class t ON t.oid = c.conrelid
                JOIN pg_namespace n ON n.oid = t.relnamespace
                JOIN pg_attribute a ON a.attrelid = t.oid AND a.attnum = ANY(c.conkey)
                WHERE t.relname = '{escape_postgresql_string(safe_table_name.strip('"'))}'
                AND n.nspname = 'public'
                AND c.contype = 'u'
                ORDER BY c.conname, a.attnum
                LIMIT 10
            """
            unique_columns = self.pg_sql_execute(unique_sql, fetch_all=True)
            
            if unique_columns:
                key_columns = [row[0] for row in unique_columns]
                key_column_names = [self.sanitise_token_for_postgresql(col) for col in key_columns]
                logger.info(f"_sync_deleted_table: {table_name} using UNIQUE constraint: {', '.join(key_columns)}")
            else:
                # No key available - cannot sync deletions
                print(f"      ⚠ WARNING: {table_name} has no PRIMARY KEY or UNIQUE constraint - cannot sync deletions")
                logger.warning(f"_sync_deleted_table: {table_name} has no PK or UNIQUE constraint - skipping")
                # Store validation result with failure
                self.validation_results[table_name] = {
                    'source_count': access_count,
                    'target_count': pg_count,
                    'pg_count_before': pg_count,
                    'matched': False,
                    'difference': access_count - pg_count,
                    'deleted': 0,
                    'error': 'No PRIMARY KEY or UNIQUE constraint'
                }
                return
        
        # Step 2: Get all column names for SELECT
        columns_sql = f"""
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema = 'public'
            AND LOWER(table_name) = LOWER('{escape_postgresql_string(safe_table_name.strip('"'))}')
            ORDER BY ordinal_position
        """
        all_columns = self.pg_sql_execute(columns_sql, fetch_all=True)
        all_column_names = [row[0] for row in (all_columns or [])]
        
        if not all_column_names:
            print(f"      ⚠ WARNING: {table_name} has no columns? - skipping")
            logger.error(f"_sync_deleted_table: {table_name} has no columns - skipping")
            self.validation_results[table_name] = {
                'source_count': access_count,
                'target_count': pg_count,
                'pg_count_before': pg_count,
                'matched': False,
                'difference': access_count - pg_count,
                'deleted': 0,
                'error': 'Table has no columns'
            }
            return
        
        # Build ORDER BY clause for pagination (must be deterministic)
        # Important: Use sanitise_token_for_postgresql to handle spaces, mixed case, and reserved words
        quoted_key_columns = [self.sanitise_token_for_postgresql(col) for col in key_columns]
        order_by_clause = f"ORDER BY {', '.join(quoted_key_columns)}"
        
        # Build WHERE clause template for Access lookup
        access_where_parts = []
        for key_col in key_columns:
            access_where_parts.append(f"[{key_col}] = ?")
        access_where_clause = " AND ".join(access_where_parts)
        
        # Batch processing parameters
        batch_size = 1000
        total_deleted = 0
        total_errors = 0
        last_key_values = None
        batch_number = 0
        check_start_time = time.time()
        last_progress_time = time.time()
        progress_interval = 2
        
        print(f"      Processing in batches of {batch_size} rows...")
        
        while True:
            batch_number += 1
            
            # Build the SELECT query with pagination
            if last_key_values is None:
                # First batch - no WHERE clause
                select_sql = f"""
                    SELECT {', '.join([self.sanitise_token_for_postgresql(col) for col in all_column_names])}
                    FROM {safe_table_name}
                    {order_by_clause}
                    LIMIT {batch_size}
                """
                params = None
            else:
                # Subsequent batches - use WHERE with > last key values
                # Validate that last_key_values has the correct length
                if len(last_key_values) != len(key_columns):
                    logger.error(f"_sync_deleted_table: last_key_values length {len(last_key_values)} does not match key_columns length {len(key_columns)} for {table_name}")
                    break
                
                if len(key_columns) == 1:
                    # Single key column - simple comparison with quoted column name
                    quoted_key_col = self.sanitise_token_for_postgresql(key_columns[0])
                    select_sql = f"""
                        SELECT {', '.join([self.sanitise_token_for_postgresql(col) for col in all_column_names])}
                        FROM {safe_table_name}
                        WHERE {quoted_key_col} > %s
                        {order_by_clause}
                        LIMIT {batch_size}
                    """
                    params = [last_key_values[0]] if last_key_values else None
                else:
                    # For composite keys, use tuple comparison with quoted column names
                    quoted_key_columns = [self.sanitise_token_for_postgresql(col) for col in key_columns]
                    placeholders = ', '.join(['%s'] * len(key_columns))
                    select_sql = f"""
                        SELECT {', '.join([self.sanitise_token_for_postgresql(col) for col in all_column_names])}
                        FROM {safe_table_name}
                        WHERE ({', '.join(quoted_key_columns)}) > ({placeholders})
                        {order_by_clause}
                        LIMIT {batch_size}
                    """
                    params = list(last_key_values) if last_key_values else None
            
            # Execute the SELECT for this batch
            try:
                if params is None:
                    self.pg_cursor.execute(select_sql)
                else:
                    self.pg_cursor.execute(select_sql, params)
                batch_rows = self.pg_cursor.fetchall()
            except Exception as e:
                print(f"      ✗ ERROR: Failed to fetch batch {batch_number} from PostgreSQL for {table_name}: {e}")
                logger.error(f"_sync_deleted_table: Failed to fetch batch {batch_number} from {table_name}: {e}")
                break
            
            if not batch_rows:
                # No more rows to process
                break
            
            # Process this batch: check existence in Access and collect keys to delete
            keys_to_delete = []
            
            # Map column names to indices in the result row
            key_indices = []
            for key_col in key_columns:
                for idx, col_name in enumerate(all_column_names):
                    if col_name.lower() == key_col.lower():
                        key_indices.append(idx)
                        break
            
            if len(key_indices) != len(key_columns):
                print(f"      ⚠ WARNING: Could not locate all key columns in {table_name}")
                logger.warning(f"_sync_deleted_table: Key columns {key_columns} not found in {table_name}")
                break
            
            # For each row in the batch, check if it exists in Access
            for pg_row in batch_rows:
                # Extract key values with bounds checking
                current_key_values = []
                valid_row = True
                for idx in key_indices:
                    if idx >= len(pg_row):
                        logger.warning(f"_sync_deleted_table: Row has {len(pg_row)} columns but index {idx} requested for {table_name}")
                        valid_row = False
                        break
                    current_key_values.append(pg_row[idx])
                
                if not valid_row:
                    continue
                
                current_key_tuple = tuple(current_key_values)
                
                # Build Access SQL to check existence
                access_key_values = []
                for idx in key_indices:
                    if idx >= len(pg_row):
                        val = None
                    else:
                        val = pg_row[idx]
                    if val is None:
                        access_key_values.append("NULL")
                    elif isinstance(val, str):
                        escaped = val.replace("'", "''")
                        access_key_values.append(f"'{escaped}'")
                    elif isinstance(val, (int, float)):
                        access_key_values.append(str(val))
                    elif isinstance(val, datetime):
                        access_key_values.append(f"#{val.strftime('%Y-%m-%d %H:%M:%S')}#")
                    elif isinstance(val, date):
                        access_key_values.append(f"#{val.strftime('%Y-%m-%d')}#")
                    elif isinstance(val, bool):
                        access_key_values.append("True" if val else "False")
                    else:
                        access_key_values.append(f"'{escape_postgresql_string(str(val))}'")
                
                access_sql = f"SELECT COUNT(*) FROM [{table_name}] WHERE {access_where_clause}"
                for i, kv in enumerate(access_key_values):
                    access_sql = access_sql.replace("?", kv, 1)
                
                try:
                    recordset = self.dao_conn.OpenRecordset(access_sql)
                    count = recordset.Fields(0).value
                    recordset.Close()
                    
                    if count == 0:
                        keys_to_delete.append(current_key_tuple)
                except Exception as e:
                    logger.warning(f"_sync_deleted_table: Error checking row in Access for {table_name}: {e}")
                    continue
            
            # Delete the collected keys for this batch
            if keys_to_delete:
                # Delete in sub-batches if needed (though batch is already limited to batch_size)
                delete_batch_size = 1000
                for i in range(0, len(keys_to_delete), delete_batch_size):
                    sub_batch = keys_to_delete[i:i+delete_batch_size]
                    
                    # Build parameterised DELETE query with quoted key columns
                    quoted_key_columns = [self.sanitise_token_for_postgresql(col) for col in key_columns]
                    row_placeholders = []
                    flat_values = []
                    for pk_tuple in sub_batch:
                        row_placeholders.append(f"({','.join(['%s'] * len(key_columns))})")
                        flat_values.extend(pk_tuple)
                    
                    delete_sql = f"""
                        DELETE FROM {safe_table_name}
                        WHERE ({', '.join(quoted_key_columns)}) IN ({', '.join(row_placeholders)})
                    """
                    
                    try:
                        self.pg_cursor.execute(delete_sql, flat_values)
                        total_deleted += len(sub_batch)
                    except Exception as e:
                        logger.error(f"_sync_deleted_table: Failed to delete batch from {table_name}: {e}")
                        logger.debug(f"Failed DELETE SQL: {delete_sql[:500]}")
                        total_errors += len(sub_batch)
                
                self.pg_conn.commit()
            
            # Remember the last key value for the next batch with bounds checking
            if batch_rows and key_indices:
                try:
                    last_row = batch_rows[-1]
                    # Verify row has expected number of columns
                    max_idx_needed = max(key_indices) if key_indices else -1
                    if max_idx_needed >= 0 and len(last_row) > max_idx_needed:
                        last_key_values = tuple(last_row[idx] for idx in key_indices)
                    else:
                        logger.warning(f"_sync_deleted_table: Row has {len(last_row)} columns but key_indices require index up to {max_idx_needed} for {table_name}")
                        last_key_values = None
                except IndexError as e:
                    logger.error(f"_sync_deleted_table: Index error extracting keys for {table_name} batch {batch_number}: {e}")
                    last_key_values = None
            else:
                last_key_values = None
                if batch_rows:
                    logger.warning(f"_sync_deleted_table: Could not determine last_key_values for {table_name} batch {batch_number}")
            
            # Progress indication with ETA
            current_time = time.time()
            if current_time - last_progress_time >= progress_interval:
                # Estimate total rows processed (cumulative)
                rows_processed_approx = batch_number * batch_size
                if access_count > 0:
                    percentage = min(100, (rows_processed_approx / access_count * 100)) if access_count > 0 else 0
                    elapsed = current_time - check_start_time
                    eta_str = ""
                    if elapsed > 0 and rows_processed_approx > 0 and rows_processed_approx < access_count:
                        rate = rows_processed_approx / elapsed
                        remaining_seconds = (access_count - rows_processed_approx) / rate
                        eta_str = f" ETA: {self.format_eta(remaining_seconds)}"
                    
                    bar_length = 40
                    filled_length = int(bar_length * rows_processed_approx // access_count) if access_count > 0 else 0
                    bar = '█' * filled_length + '░' * (bar_length - filled_length)
                    
                    print(f"      Progress: Batch {batch_number}, Processed approx {rows_processed_approx}/{access_count} rows [{bar}] {percentage:.1f}%{eta_str} - Deleted {total_deleted} so far", end='\r', flush=True)
                last_progress_time = current_time
        
        print()  # Newline after progress
        
        self.pg_conn.commit()
        
        # Re-query final count
        count_sql = f"SELECT COUNT(*) FROM {safe_table_name}"
        result = self.pg_sql_execute(count_sql, fetch_one=True)
        final_pg_count = result[0] if result else 0
        
        # Calculate total elapsed time
        total_elapsed = time.time() - sync_start_time
        elapsed_str = self.format_elapsed(total_elapsed)
        
        print(f"      ✓ Completed {table_name}: Deleted {total_deleted} rows (errors: {total_errors}) (elapsed: {elapsed_str})")
        logger.info(f"_sync_deleted_table: Completed {table_name} - deleted {total_deleted}, errors {total_errors}, elapsed {total_elapsed:.2f}s")
        
        # Update validation results after deletion with before/after counts
        self.validation_results[table_name] = {
            'source_count': access_count,
            'target_count': final_pg_count,
            'pg_count_before': original_pg_count,
            'matched': access_count == final_pg_count,
            'difference': access_count - final_pg_count,
            'deleted': total_deleted,
            'errors': total_errors
        }
    
    # ========================================================================
    # LIST DATABASE TABLES (DAO ONLY - NO POSTGRESQL)
    # ========================================================================
    
    def list_database_tables(self):
        """List all non-system tables from MS Access database only."""
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - list_database_tables called")
        
        logger.info("====== listing database tables ======")
        
        # Open only DAO connection (no PostgreSQL needed)
        try:
            self.open_DAO_connection()
        except Exception as e:
            print(f"Error: Failed to connect to MS Access database: {e}")
            logger.error(f"list_database_tables: DAO connection failed - {e}")
            return
        
        tables_to_process = self.get_all_tables_to_process()
        
        print("\nTables in database:")
        print("-" * 40)
        for table_name in tables_to_process:
            print(f"  {table_name}")
        print("-" * 40)
        print(f"Total: {len(tables_to_process)} tables")
        
        logger.info(f"Total tables listed: {len(tables_to_process)}")
        
        # Close only DAO connection
        if self.dao_conn:
            self.dao_conn.Close()
            self.dao_conn = None
            logger.info("DAO connection closed")
        
        logger.info("====== list completed =====")
    
    # ========================================================================
    # FOREIGN KEY DISCOVERY
    # ========================================================================
    # IMPORTANT TERMINOLOGY NOTE:
    # MS Access stores relationships where:
    #   rel.Table        = the CHILD table (the one with the foreign key)
    #   rel.ForeignTable = the PARENT table (the referenced table)
    # However, in this program we store them as:
    #   'base_table'      = child (FK side)
    #   'reference_table' = parent (referenced side)
    # This naming is inverted from standard SQL terminology, but we keep it
    # for backward compatibility. When reading the code, remember:
    #   base_table     -> child table (where FK constraint is defined)
    #   reference_table -> parent table (referenced by the FK)
    # ========================================================================
    
    def get_foreign_keys(self):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - get_foreign_keys called")
        
        self.original_configured_fkeys = self.parameters.get('foreignkeys', {}).copy()
        discovered_fkeys = {}
        
        for i in range(self.dao_conn.Relations.Count):
            rel = self.dao_conn.Relations[i]
            rel_name = decode_sketchy_utf16(rel.Name)
            
            if self.debug:
                logger.debug(f"Processing Relation: {rel_name}")
                if hasattr(rel, 'Attributes'):
                    logger.debug(f"  Attributes: {rel.Attributes}")
            
            if rel.Fields.Count == 0:
                logger.error(f"FATAL: Relation {rel_name} has Fields.Count = 0")
                logger.error(f"  Relation name: '{rel_name}'")
                logger.error(f"  This indicates a corrupted or empty relation in MS Access")
                logger.error(f"  Please check the MS Access database relationships")
                sys.exit(1)
            
            # MS Access: rel.Table = child (FK side), rel.ForeignTable = parent (referenced)
            child_table = decode_sketchy_utf16(rel.Table)
            parent_table = decode_sketchy_utf16(rel.ForeignTable)
            child_columns = []
            parent_columns = []
            
            for j in range(rel.Fields.Count):
                field = rel.Fields[j]
                child_columns.append(decode_sketchy_utf16(field.Name))
                parent_columns.append(decode_sketchy_utf16(field.ForeignName))
            
            fk_name = f"fk_{child_table}_{parent_table}"
            fk_key = fk_name.lower()
            
            # Store with internal names: base_table = child, reference_table = parent
            discovered_fkeys[fk_key] = {
                'base_table': child_table,
                'reference_table': parent_table,
                'base_columns': child_columns,
                'reference_columns': parent_columns
            }
            
            logger.debug(f"Discovered foreign key: {fk_key}")
            if self.debug:
                logger.debug(f"  {child_table}({child_columns}) -> {parent_table}({parent_columns})")
        
        # Filter discovered foreign keys - handle excluded tables and system tables
        filtered_fkeys = {}
        for fk_name, fk_info in discovered_fkeys.items():
            base_table = fk_info['base_table']      # child
            reference_table = fk_info['reference_table']  # parent
            
            # Fatal error if the child table (base_table) is excluded
            if self.is_table_excluded(base_table):
                logger.error(f"FATAL: Foreign key '{fk_name}' has base_table '{base_table}' in excluded list")
                logger.error(f"  Cannot create foreign key on excluded table")
                sys.exit(1)
            
            # Skip if child table is a system table
            if self.is_system_table_name(base_table):
                logger.warning(f"Skipping foreign key '{fk_name}' - base_table '{base_table}' is a system table")
                continue
            
            # Skip if parent table is a system table
            if self.is_system_table_name(reference_table):
                logger.warning(f"Skipping foreign key '{fk_name}' - reference_table '{reference_table}' is a system table")
                continue
            
            # Skip (non-fatal) if only the parent table is excluded
            if self.is_table_excluded(reference_table):
                logger.warning(f"Skipping foreign key '{fk_name}' - reference_table '{reference_table}' is excluded")
                continue
            
            # Keep this foreign key
            filtered_fkeys[fk_name] = fk_info
        
        self.discovered_foreign_keys = filtered_fkeys.copy()
        
        merged_fkeys = filtered_fkeys.copy()
        merged_fkeys.update(self.original_configured_fkeys)
        self.parameters['foreignkeys'] = merged_fkeys
        
        for fk_name, fk_info in self.original_configured_fkeys.items():
            if fk_name in filtered_fkeys:
                discovered_info = filtered_fkeys[fk_name]
                if (discovered_info['base_table'] != fk_info['base_table'] or
                    discovered_info['reference_table'] != fk_info['reference_table'] or
                    discovered_info['base_columns'] != fk_info['base_columns'] or
                    discovered_info['reference_columns'] != fk_info['reference_columns']):
                    logger.error(f"FATAL: Foreign key name collision with different definition")
                    logger.error(f"  Foreign key name: {fk_name}")
                    logger.error(f"  Configured: {fk_info['base_table']}({fk_info['base_columns']}) -> {fk_info['reference_table']}({fk_info['reference_columns']})")
                    logger.error(f"  Discovered: {discovered_info['base_table']}({discovered_info['base_columns']}) -> {discovered_info['reference_table']}({discovered_info['reference_columns']})")
                    sys.exit(1)
        
        logger.info(f"Discovered {len(filtered_fkeys)} foreign keys from MS Access (after exclusion and system table filtering)")
        logger.info(f"Configured {len(self.original_configured_fkeys)} foreign keys from YAML")
        logger.info(f"Total foreign keys: {len(self.parameters['foreignkeys'])}")
        
        return self.parameters['foreignkeys']
    
    # ========================================================================
    # CREATE TABLE
    # ========================================================================
    
    def create_primary_key_clause(self, table_info):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - create_primary_key_clause called for table: {table_info.get('name', 'unknown')}")
        
        pk_columns = []
        for idx in table_info.get('indexes', []):
            if idx.get('primary'):
                for col_name in idx['columns']:
                    for col in table_info['columns']:
                        if col['original_name'] == col_name or \
                           self.normalize_column_name(col['original_name']) == self.normalize_column_name(col_name):
                            pk_columns.append(col['name'])
                            break
                break
        if pk_columns:
            return f"    PRIMARY KEY ({', '.join(pk_columns)})"
        return ''
    
    def create_table(self, table_info):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - create_table called for table: {table_info.get('name', 'unknown')}")
        
        safe_table_name = table_info['safe_name']
        
        table_exists = self.table_exists_in_postgresql(safe_table_name)
        
        if table_exists:
            msg = f"Table already exists: {table_info['name']}"
            logger.info(msg)
            print(msg)
            return
        
        columns_def = []
        for col in table_info['columns']:
            col_def = f"    {col['name']} {col['type']}"
            if col.get('required'):
                col_def += " NOT NULL"
            columns_def.append(col_def)
        
        pk_clause = self.create_primary_key_clause(table_info)
        if pk_clause:
            create_sql = f"CREATE TABLE {safe_table_name} (\n{',\n'.join(columns_def)},\n{pk_clause}\n)"
        else:
            create_sql = f"CREATE TABLE {safe_table_name} (\n{',\n'.join(columns_def)}\n)"
        
        self.pg_sql_execute(create_sql)
        msg = f"Created table: {table_info['name']}"
        logger.info(msg)
        print(msg)
    
    def create_all_target_tables(self, tables_to_process):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - create_all_target_tables called with {len(tables_to_process)} tables")
        
        for table_info in tables_to_process:
            self.create_table(table_info)
            self.pg_conn.commit()
    
    # ========================================================================
    # CREATE INDEXES
    # ========================================================================
    
    def create_all_indexes(self, tables_to_process):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - create_all_indexes called with {len(tables_to_process)} tables")
        
        for table_info in tables_to_process:
            safe_table_name = table_info['safe_name']
            for idx in table_info.get('indexes', []):
                if idx.get('primary'):
                    continue
                
                idx_columns = []
                for col_name in idx['columns']:
                    for col in table_info['columns']:
                        if col['original_name'] == col_name or \
                           self.normalize_column_name(col['original_name']) == self.normalize_column_name(col_name):
                            idx_columns.append(col['name'])
                            break
                
                if not idx_columns:
                    continue
                
                safe_idx_name = self.normalise_index_name(table_info['name'], idx['name'])
                
                check_sql = f"""
                    SELECT indexname FROM pg_indexes 
                    WHERE LOWER(tablename) = LOWER('{escape_postgresql_string(table_info['name'])}')
                    AND LOWER(indexname) = LOWER('{escape_postgresql_string(safe_idx_name)}')
                    AND schemaname = 'public'
                """
                exists = self.pg_sql_execute(check_sql, fetch_one=True)
                
                if exists:
                    logger.info(f"Index {idx['name']} on table {table_info['name']} already exists - skipping")
                    continue
                
                if idx.get('unique'):
                    sql = f"CREATE UNIQUE INDEX {safe_idx_name} ON {safe_table_name} ({', '.join(idx_columns)})"
                else:
                    sql = f"CREATE INDEX {safe_idx_name} ON {safe_table_name} ({', '.join(idx_columns)})"
                
                self.pg_sql_execute(sql)
                logger.info(f"Created index: {idx['name']} on {table_info['name']}")
                self.pg_conn.commit()
    
    # ========================================================================
    # CREATE FOREIGN KEYS (WITH AUTO-CREATION OF UNIQUENESS ON BASE TABLE)
    # ========================================================================
    # IMPORTANT: In this program, 'base_table' is the child (FK side) and
    # 'reference_table' is the parent. The foreign key is created on the
    # base_table (child) referencing the reference_table (parent).
    # ========================================================================
    
    def create_foreign_key(self, fk_name, fk_info):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - create_foreign_key called with params: fk_name={fk_name}")
        
        base_table = fk_info['base_table']          # child table (where FK is defined)
        reference_table = fk_info['reference_table']  # parent table (referenced)
        base_columns = fk_info['base_columns']
        reference_columns = fk_info['reference_columns']
        
        is_from_config = fk_name in self.original_configured_fkeys
        
        corrected = self.attempt_direction_correction(fk_name, fk_info, is_from_config)
        
        if corrected is not None:
            self.parameters['foreignkeys'][fk_name] = corrected.copy()
            if fk_name in self.discovered_foreign_keys:
                self.discovered_foreign_keys[fk_name] = corrected.copy()
            self.create_foreign_key(fk_name, corrected)
            return
        elif corrected is None and self.is_direction_valid(base_table, base_columns, reference_table, reference_columns) is None:
            # Neither table has uniqueness - normally we would auto-create a constraint,
            # but if --no-auto-index is set, we skip this foreign key entirely.
            if self.no_auto_index:
                # Build warning banner
                warning_banner = f"""
{'='*70}
⚠️  FOREIGN KEY SKIPPED (--no-auto-index enabled):
    Foreign key name: {fk_name}
    Child table (base_table)  : {base_table}
    Parent table (reference)   : {reference_table}
    Child columns              : {', '.join(base_columns)}
    Parent columns             : {', '.join(reference_columns)}
    
    This foreign key requires a PRIMARY KEY or UNIQUE constraint on the child
    table columns, but automatic creation is disabled.
    The foreign key will NOT be created.
{'='*70}
"""
                print(warning_banner)
                logger.warning(warning_banner.replace('\n', ' ').strip())
                return  # skip this FK
            else:
                # Auto-creation is allowed (default behaviour)
                logger.info(f"Foreign key '{fk_name}': base table '{base_table}' lacks uniqueness on columns {base_columns}")
                if self.is_system_table_name(base_table):
                    logger.warning(f"Cannot add uniqueness constraint to system table '{base_table}' - skipping foreign key")
                    if is_from_config:
                        logger.error(f"FATAL: Foreign key from CONFIGURATION FILE references system table")
                        sys.exit(1)
                    return
                try:
                    constraint_name = self.ensure_uniqueness_on_base_table(base_table, base_columns, fk_name)
                    logger.info(f"Retrying foreign key creation after ensuring uniqueness on base table")
                    self.create_foreign_key(fk_name, fk_info)
                    return
                except Exception as e:
                    logger.error(f"Failed to automatically add uniqueness constraint for foreign key '{fk_name}': {e}")
                    if is_from_config:
                        logger.error(f"FATAL: Foreign key from CONFIGURATION FILE cannot be created")
                        sys.exit(1)
                    else:
                        logger.warning(f"Skipping foreign key from MS Access discovery")
                        return
        
        # Safety check - skip if reference table is excluded (should already be filtered, but check anyway)
        if self.is_table_excluded(reference_table):
            logger.info(f"Skipping foreign key '{fk_name}' - reference_table '{reference_table}' is excluded")
            return
        
        # Skip if base_table is a system table
        if self.is_system_table_name(base_table):
            logger.info(f"Skipping foreign key '{fk_name}' - base_table '{base_table}' is a system table")
            return
        
        # Skip if reference_table is a system table
        if self.is_system_table_name(reference_table):
            logger.info(f"Skipping foreign key '{fk_name}' - reference_table '{reference_table}' is a system table")
            return
        
        safe_fk_name = self.sanitise_keyname_for_postgresql(fk_name)
        safe_base_table = self.sanitise_token_for_postgresql(base_table)
        safe_ref_table = self.sanitise_token_for_postgresql(reference_table)
        
        # IMPROVED CHECK - Check specifically on the base table where the constraint will be created
        check_sql = f"""
            SELECT 1
            FROM information_schema.table_constraints 
            WHERE constraint_type = 'FOREIGN KEY'
            AND LOWER(constraint_name) = LOWER('{escape_postgresql_string(safe_fk_name)}')
            AND LOWER(table_name) = LOWER('{escape_postgresql_string(base_table)}')
            AND table_schema = 'public'
        """
        existing = self.pg_sql_execute(check_sql, fetch_one=True)
        
        if existing:
            logger.info(f"Foreign key {fk_name} already exists on table {base_table} - skipping creation")
            if self.verbose:
                print(f"  Foreign key {fk_name} already exists - skipping")
            return
        
        base_mapping = self.parameters.get('column_mapping', {}).get(base_table, {})
        ref_mapping = self.parameters.get('column_mapping', {}).get(reference_table, {})
        
        resolved_base = []
        resolved_ref = []
        for col in base_columns:
            normalized = self.normalize_column_name(col)
            resolved_base.append(base_mapping.get(normalized, self.sanitise_token_for_postgresql(col)))
        for col in reference_columns:
            normalized = self.normalize_column_name(col)
            resolved_ref.append(ref_mapping.get(normalized, self.sanitise_token_for_postgresql(col)))
        
        safe_fk_name = self.sanitise_keyname_for_postgresql(fk_name)
        safe_base_table = self.sanitise_token_for_postgresql(base_table)
        safe_ref_table = self.sanitise_token_for_postgresql(reference_table)
        
        alter_sql = f"""
            ALTER TABLE {safe_base_table} 
            ADD CONSTRAINT {safe_fk_name} 
            FOREIGN KEY ({', '.join(resolved_base)}) 
            REFERENCES {safe_ref_table} ({', '.join(resolved_ref)})
            ON DELETE CASCADE
        """
        
        try:
            self.pg_sql_execute(alter_sql)
            logger.info(f"Created foreign key: {fk_name}")
            self.pg_conn.commit()
        except Exception as e:
            self.pg_conn.rollback()
            logger.error(f"Failed to create foreign key {fk_name}: {e}")
            if is_from_config:
                sys.exit(1)
            else:
                logger.warning(f"Skipping discovered foreign key")
    
    def create_all_foreign_keys(self):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - create_all_foreign_keys called")
        
        fkeys = self.parameters.get('foreignkeys', {})
        skipped_count = 0
        
        for fk_name, fk_info in fkeys.items():
            base_table = fk_info.get('base_table', '')
            reference_table = fk_info.get('reference_table', '')
            
            # Skip if base_table is a system table
            if self.is_system_table_name(base_table):
                logger.info(f"Skipping foreign key '{fk_name}' - base_table '{base_table}' is a system table")
                skipped_count += 1
                continue
            
            # Skip if reference_table is a system table
            if self.is_system_table_name(reference_table):
                logger.info(f"Skipping foreign key '{fk_name}' - reference_table '{reference_table}' is a system table")
                skipped_count += 1
                continue
            
            self.create_foreign_key(fk_name, fk_info)
            self.pg_conn.commit()
        
        if skipped_count > 0:
            logger.info(f"Skipped {skipped_count} foreign key(s) involving system tables")
    
    # ========================================================================
    # COPY ALL TABLES WITH DEPENDENCY RESOLUTION
    # ========================================================================
    
    def copy_all_base_tables(self, tables_to_process):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - copy_all_base_tables called")
        
        print("\n" + "=" * 60)
        print("COPYING TABLES WITHOUT FOREIGN KEYS")
        print("=" * 60)
        
        self.copied_tables = set()
        fkey_tables = {fk_info.get('base_table') for fk_info in self.parameters.get('foreignkeys', {}).values()}
        
        for table_info in tables_to_process:
            if table_info['name'] not in fkey_tables:
                self.copy_table(table_info)
                self.copied_tables.add(table_info['name'])
        logger.info("====== end of copying base tables =====")
    
    def copy_all_fkey_tables(self, tables_to_process):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - copy_all_fkey_tables called")
        
        print("\n" + "=" * 60)
        print("COPYING TABLES WITH FOREIGN KEYS")
        print("=" * 60)
        
        copied_tables = self.copied_tables.copy()
        fk_config = self.parameters.get('foreignkeys', {})
        fkey_tables = {fk_info.get('base_table') for fk_info in fk_config.values()}
        
        dependencies = {}
        for fk_info in fk_config.values():
            child = fk_info.get('base_table', '')
            parent = fk_info.get('reference_table', '')
            if child not in dependencies:
                dependencies[child] = set()
            dependencies[child].add(parent)
        
        MAX_ITERATIONS = 25
        for iteration in range(MAX_ITERATIONS):
            self.current_fk_iteration = iteration + 1
            copied_this_round = 0
            for table_info in tables_to_process:
                table_name = table_info['name']
                if table_name not in fkey_tables or table_name in copied_tables:
                    continue
                deps = dependencies.get(table_name, set())
                if deps and not deps.issubset(copied_tables):
                    continue
                self.copy_table(table_info)
                copied_tables.add(table_name)
                copied_this_round += 1
            if copied_this_round == 0:
                remaining = [t for t in fkey_tables if t not in copied_tables]
                if remaining:
                    logger.warning(f"Could not copy remaining FK tables: {remaining}")
                break
        
        self.current_fk_iteration = 0
        logger.info("====== end of copying fkey tables =====")
    
    # ========================================================================
    # REPLICATION METHODS
    # ========================================================================
    
    def replicate_schema(self):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - replicate_schema called")
        
        pg_params = self.parameters.get('postgresql', {})
        dbname = pg_params.get('database', '')
        
        admin_params = pg_params.copy()
        admin_params['database'] = 'postgres'
        admin_conn = psycopg2.connect(
            host=admin_params['host'],
            port=int(admin_params['port']),
            database='postgres',
            user=admin_params['user'],
            password=admin_params['password']
        )
        admin_conn.autocommit = True
        admin_cursor = admin_conn.cursor()
        
        admin_cursor.execute(f"DROP DATABASE IF EXISTS {dbname}")
        logger.info(f"Dropped database: {dbname}")
        admin_cursor.execute(f"CREATE DATABASE {dbname}")
        logger.info(f"Created database: {dbname}")
        
        admin_cursor.close()
        admin_conn.close()
        
        self.open_connections()
        self.get_foreign_keys()
        table_names = self.get_all_tables_to_process()
        
        tables_to_process = []
        for table_name in table_names:
            tables_to_process.append(self.load_table(table_name))
        
        self.tables_info = tables_to_process
        self.create_all_target_tables(tables_to_process)
        self.create_all_indexes(tables_to_process)
        self.create_all_foreign_keys()
        
        logger.info("Schema replication only - data copy skipped")
        if self.verbose:
            print("Schema replication only - data copy skipped")
        
        self.close_connections()
    
    def replicate_tables(self):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - replicate_tables called")
        
        # Record program start time for total elapsed calculation
        self.program_start_time = time.time()
        
        self.validate_configuration()
        
        self.open_connections()
        if not self.validate_transformation():
            logger.error("invalid transformations - check your spelling, capitals and spaces")
            logger.error("FATAL: invalid transformations - no point in continuing")
            print("\ninvalid transformations - check your spelling, capitals and spaces")
            print("\nFATAL: invalid transformations - no point in continuing")
            self.close_connections()
            sys.exit(1)

        self.get_foreign_keys()
        table_names = self.get_all_tables_to_process()
        tables_to_process = []
        for table_name in table_names:
            tables_to_process.append(self.load_table(table_name))
        
        self.tables_info = tables_to_process
        
        if self.full_refresh:
            print("\nFULL REFRESH mode - dropping all tables...")
            for table_info in reversed(tables_to_process):
                safe_name = table_info['safe_name']
                drop_sql = f"DROP TABLE IF EXISTS {safe_name} CASCADE"
                self.pg_sql_execute(drop_sql)
                self.pg_conn.commit()
            print("Dropped all existing tables")

        print("\n" + "=" * 60)
        print("CREATING TABLES AND INDEXES")
        print("=" * 60)
        
        self.create_all_target_tables(tables_to_process)
        self.create_all_indexes(tables_to_process)
        
        logger.info("Creating foreign keys...")
        self.create_all_foreign_keys()
        
        logger.info("Copying data...")
        self.copy_all_base_tables(tables_to_process)
        self.copy_all_fkey_tables(tables_to_process)
        
        if self.sync_deleted:
            self.sync_deleted_tables(tables_to_process)
        
        self.print_validation_summary()
        
        # Calculate and display total elapsed time
        if self.program_start_time:
            total_elapsed = time.time() - self.program_start_time
            total_elapsed_str = self.format_elapsed(total_elapsed)
            msg = f"\nTotal replication time: {total_elapsed_str}"
            logger.info(msg)
            print(msg)
        
        self.close_connections()
    
    def test_network_connections(self):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - test_network_connections called")
        
        logger.info("====== testing network connections ======")
        try:
            self.open_DAO_connection()
            logger.info("MS Access connection: SUCCESS")
            if self.verbose:
                print("MS Access connection: SUCCESS")
            self.dao_conn.Close()
        except Exception as e:
            logger.error(f"MS Access connection: FAILED - {e}")
            if self.verbose:
                print(f"MS Access connection: FAILED - {e}")
        
        try:
            # Use postgres database for connection test
            self.open_postgresql_connection_master()
            logger.info("PostgreSQL connection: SUCCESS")
            if self.verbose:
                print("PostgreSQL connection: SUCCESS")
            self.pg_conn.close()
        except Exception as e:
            logger.error(f"PostgreSQL connection: FAILED - {e}")
            if self.verbose:
                print(f"PostgreSQL connection: FAILED - {e}")
        
        logger.info("====== network test completed =====")
    
    def dump_internal_data(self):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - dump_internal_data called")
        
        print("\n" + "=" * 60)
        print("INTERNAL PROGRAM DATA DUMP")
        print("=" * 60)
        
        print("\n--- PARAMETERS ---")
        for key in self.parameters:
            if key == 'postgresql':
                print(f"  {key}: {{'host': '***', 'port': '***', 'database': '***', 'user': '***', 'password': '***'}}")
            else:
                print(f"  {key}: {self.parameters[key]}")
        
        print("\n--- COPIED TABLES ---")
        print(f"  {self.copied_tables}")
        
        print("\n--- ORIGINAL CONFIGURED FKEYS ---")
        for key in self.original_configured_fkeys:
            print(f"  {key}: {self.original_configured_fkeys[key]}")
        
        print("\n--- DISCOVERED FOREIGN KEYS ---")
        for key in self.discovered_foreign_keys:
            print(f"  {key}: {self.discovered_foreign_keys[key]}")
        
        print("\n--- FLAGS ---")
        print(f"  trace: {self.trace}")
        print(f"  verbose: {self.verbose}")
        print(f"  debug: {self.debug}")
        print(f"  list_tables: {self.list_tables}")
        print(f"  test_network: {self.test_network}")
        print(f"  dump_data: {self.dump_data}")
        print(f"  schema_only: {self.schema_only}")
        print(f"  full_refresh: {self.full_refresh}")
        print(f"  current_fk_iteration: {self.current_fk_iteration}")
        print(f"  no_auto_index: {self.no_auto_index}")
        print(f"  adjust_ms_access: {self.adjust_ms_access}")
        print(f"  sync_deleted: {self.sync_deleted}")
        print(f"  slow: {self.slow}")
        print(f"  nonvolatile: {self.nonvolatile}")
        
        print("\n--- VALIDATION RESULTS ---")
        if self.validation_results:
            for table_name, result in self.validation_results.items():
                status = "✓" if result.get('matched', False) else "✗"
                source = result.get('source_count', 'N/A')
                target = result.get('target_count', 'N/A')
                pg_before = result.get('pg_count_before', 'N/A')
                deleted = result.get('deleted', 0)
                nonvolatile_skipped = result.get('nonvolatile_skipped', False)
                nv_flag = " [NV SKIPPED]" if nonvolatile_skipped else ""
                print(f"  {status} {table_name}: Access={source}, PG Before={pg_before}, PG After={target}, Deleted={deleted}{nv_flag}")
        else:
            print("  No validation results available")
        
        print("\n--- EMPTY TABLES (0 ROWS) ---")
        if self.empty_tables:
            print(f"  Count: {len(self.empty_tables)}")
            for table in self.empty_tables:
                print(f"  - {table}")
        else:
            print("  No empty tables")
        
        print("\n" + "=" * 60)
    
    def exit_and_cleanup(self):
        if self.trace:
            frame = inspect.currentframe()
            logger.info(f"Line {frame.f_lineno} - exit_and_cleanup called")
        
        try:
            self.close_connections()
        except:
            pass
        print("\nReplication processing completed - successful")

    def generate_yaml_file(self):
        frame_lineno = inspect.currentframe().f_lineno
        if self.trace:
            logger.info(f"LINE {frame_lineno} - Entering generate_yaml_file")
        
        output_file = self.parameters.get('output_file', '')
        
        if not output_file:
            logger.error("No output file name specified")
            print("Error: No output file name specified")
            return False
        
        if os.path.exists(output_file):
            logger.error(f"Output file already exists: {output_file}")
            print(f"Error: Output file already exists: {output_file}")
            return False
        
        if not self.open_DAO_connection():
            logger.error("Failed to open DAO connection")
            return False
        
        # Load all tables first (auto-discovery if no tables section)
        table_names = self.get_all_tables_to_process()
        
        # Discover all foreign keys at once
        self.get_foreign_keys()
        
        logger.info(f"Total foreign keys discovered: {len(self.parameters.get('foreignkeys', {}))}")
        
        if len(self.parameters.get('foreignkeys', {})) == 0:
            logger.warning("No foreign keys discovered in database")
        
        output_data = {}
        
        sections_order = ['global', 'DAO', 'postgresql', 'excluded', 'nonvolatile', 'tables', 'primarykeys', 'foreignkeys', 'transformations']
        
        for section in sections_order:
            if section == 'tables':
                # Always output the tables section with discovered tables
                # Preserve any existing table configuration (like 'name' dict entries) if present
                existing_tables = self.parameters.get('tables')
                if existing_tables is None:
                    existing_tables = []
                tables_list = []
                
                if existing_tables and len(existing_tables) > 0:
                    # There is an existing tables configuration - preserve it
                    for item in existing_tables:
                        if isinstance(item, dict) and 'name' in item:
                            tables_list.append(item['name'])
                        elif isinstance(item, str):
                            tables_list.append(item)
                        else:
                            tables_list.append(str(item))
                else:
                    # No existing tables configuration - use auto-discovered tables
                    tables_list = table_names
                
                output_data[section] = tables_list
                
            elif section == 'nonvolatile':
                # Preserve existing nonvolatile entries if they exist
                existing_nonvolatile = self.parameters.get('nonvolatile')
                if existing_nonvolatile is None:
                    existing_nonvolatile = []
                if existing_nonvolatile and len(existing_nonvolatile) > 0:
                    output_data[section] = existing_nonvolatile
                # If no nonvolatile section exists, don't create one (optional)
                # To always include nonvolatile section even if empty, uncomment:
                # else:
                #     output_data[section] = []
                
            elif section == 'foreignkeys':
                fk_data = {}
                for fk_name, fk_info in self.parameters.get('foreignkeys', {}).items():
                    fk_data[fk_name] = {
                        'base_table': fk_info['base_table'],
                        'reference_table': fk_info['reference_table'],
                        'base_columns': fk_info['base_columns'],
                        'reference_columns': fk_info['reference_columns']
                    }
                output_data[section] = fk_data
                
            elif section in self.parameters:
                # For all other sections, copy existing values
                output_data[section] = self.parameters[section]
            # If section not in self.parameters, don't add it to output
        
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                yaml.dump(output_data, f, default_flow_style=False, sort_keys=False, allow_unicode=True)
            
            logger.info(f"YAML configuration saved to {output_file}")
            print(f"YAML configuration saved to {output_file}")
            return True
        except Exception as e:
            logger.error(f"Failed to write YAML file: {e}")
            print(f"Error: Failed to write YAML file: {e}")
            return False


# ============================================================================
# MAIN FUNCTION
# ============================================================================

def main():
    parser = argparse.ArgumentParser(description='MS Access to PostgreSQL Replication Tool')
    parser.add_argument('-V', '--version', action='store_true', help='Show version and exit')
    parser.add_argument('-c', '--config', default='replicatorconfig.yaml', help='Path to configuration file')
    parser.add_argument('-s', '--source', help='MS Access database file name')
    parser.add_argument('--thost', help='PostgreSQL server host name or IP')
    parser.add_argument('--tport', help='PostgreSQL server port number')
    parser.add_argument('--tdatabase', help='PostgreSQL database name')
    parser.add_argument('--tuser', help='PostgreSQL user name')
    parser.add_argument('--tpassword', help='PostgreSQL password')
    parser.add_argument('-v', '--verbose', action='store_true', help='Print informational messages')
    parser.add_argument('-o', '--output', help='Output file for generated YAML configuration')
    parser.add_argument('--debug', action='store_true', help='Enable SQL debugging output')
    parser.add_argument('--trace', action='store_true', help='Enable trace logging to file')
    parser.add_argument('-a', '--no-auto-index', action='store_true', 
                        help='Suppress automatic creation of indexes/constraints for foreign keys (skip FK if needed)')
    parser.add_argument('--sync-deleted', action='store_true',
                        help='Synchronize deleted records from Access to PostgreSQL')
    parser.add_argument('--slow', action='store_true',
                        help='Use slower processing; disables nonvolatile optimization (can be used with or without --sync-deleted)')
    parser.add_argument('--nonvolatile', action='store_true',
                        help='Skip copying non-volatile tables when row counts match (unless --slow is also enabled)')
    
    # Mutually exclusive action group
    action_group = parser.add_mutually_exclusive_group()
    action_group.add_argument('-S', '--schema', action='store_true', 
                              help='Drop and recreate database, then replicate schema ONLY')
    action_group.add_argument('--adjust-ms-access', action='store_true',
                              help='Adjust MS Access schema (add AutoNumber primary key to tables without PK)')
    action_group.add_argument('-l', '--list', action='store_true', 
                              help='List table names and exit')
    action_group.add_argument('-n', '--network', action='store_true', 
                              help='Test both source and target connections')
    action_group.add_argument('--dump', action='store_true', 
                              help='Dump internal program data')
    action_group.add_argument('--full-refresh', action='store_true',
                              help='Perform full refresh (drop and recreate all tables)')
    
    args = parser.parse_args()
    
    # --slow is now allowed without --sync-deleted (it also affects nonvolatile optimization)
    # No error condition needed
    
    version_raw = __version__
    if args.version:
        print(f"MS Access to PostgreSQL Replication Tool {version_raw}")
        return

    if args.output:
        manager = ReplicationManager()
        manager.parameters['output_file'] = args.output
        
        # Load default configuration if it exists
        if os.path.exists(args.config):
            with open(args.config, 'r', encoding='utf-8') as f:
                try:
                    default_params = yaml.safe_load(f)
                    if default_params:
                        manager.parameters.update(default_params)
                except yaml.YAMLError as e:
                    logger.warning(f"Failed to parse default config file: {e}")
        
        manager.generate_yaml_file()
        return
    
    if not os.path.exists(args.config):
        logger.error(f"Configuration file not found: {args.config}")
        print(f"Error: Configuration file not found: {args.config}")
        sys.exit(1)
    
    with open(args.config, 'r', encoding='utf-8') as f:
        try:
            parameters = yaml.safe_load(f)
            # If --adjust-ms-access is set, remove PostgreSQL section (optional)
            if args.adjust_ms_access and 'postgresql' in parameters:
                del parameters['postgresql']
                logger.info("adjust-ms-access: PostgreSQL configuration ignored")
        except yaml.YAMLError as e:
            logger.error(f"Failed to parse YAML configuration: {e}")
            print(f"Error: Failed to parse YAML configuration: {e}")
            sys.exit(1)
    
    if args.source:
        if 'DAO' not in parameters:
            parameters['DAO'] = {}
        parameters['DAO']['database'] = args.source
    
    # Only process PostgreSQL parameters if NOT in adjust-ms-access mode
    if not args.adjust_ms_access:
        if args.thost:
            if 'postgresql' not in parameters:
                parameters['postgresql'] = {}
            parameters['postgresql']['host'] = args.thost
        if args.tport:
            if 'postgresql' not in parameters:
                parameters['postgresql'] = {}
            parameters['postgresql']['port'] = int(args.tport)
        if args.tdatabase:
            if 'postgresql' not in parameters:
                parameters['postgresql'] = {}
            parameters['postgresql']['database'] = args.tdatabase
        if args.tuser:
            if 'postgresql' not in parameters:
                parameters['postgresql'] = {}
            parameters['postgresql']['user'] = args.tuser
        if args.tpassword:
            if 'postgresql' not in parameters:
                parameters['postgresql'] = {}
            parameters['postgresql']['password'] = args.tpassword
    
    if args.debug:
        logger.setLevel(logging.DEBUG)
        for handler in logger.handlers:
            if isinstance(handler, logging.StreamHandler):
                handler.setLevel(logging.DEBUG)
    elif args.trace:
        logger.setLevel(logging.INFO)
    
    if args.verbose:
        for handler in logger.handlers:
            if isinstance(handler, logging.StreamHandler):
                handler.setLevel(logging.INFO)
    
    manager = ReplicationManager()
    manager.trace = args.trace
    manager.verbose = args.verbose
    manager.debug = args.debug
    manager.list_tables = args.list
    manager.test_network = args.network
    manager.dump_data = args.dump
    manager.schema_only = args.schema
    manager.full_refresh = args.full_refresh
    manager.no_auto_index = args.no_auto_index
    manager.adjust_ms_access = args.adjust_ms_access
    manager.sync_deleted = args.sync_deleted
    manager.slow = args.slow
    manager.nonvolatile = args.nonvolatile
    manager.parameters = parameters
    
    print("Replication processing starts")
    
    try:
        if manager.list_tables:
            manager.list_database_tables()  # Only DAO connection, no PostgreSQL
        elif manager.test_network:
            manager.test_network_connections()
        elif manager.dump_data:
            manager.open_connections()
            manager.get_foreign_keys()
            manager.get_all_tables_to_process()
            manager.dump_internal_data()
        elif manager.schema_only:
            manager.replicate_schema()
        elif manager.adjust_ms_access:
            # Only DAO connection needed - PostgreSQL not opened
            manager.adjust_ms_access_schema()
        else:
            # Normal replication (including --full-refresh if set)
            manager.replicate_tables()
            manager.exit_and_cleanup()
    except Exception as e:
        logger.error(f"Replication processing completed - with errors: {e}")
        if manager.verbose:
            print(f"Replication processing completed - with errors: {e}")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()