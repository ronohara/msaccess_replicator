# MS Access to PostgreSQL Data Type Mappings

## Summary Table

| MS Access Data Type | DAO Type Code | PostgreSQL Data Type | Notes |
|---------------------|---------------|----------------------|-------|
| dbBoolean | 1 | BOOLEAN | TRUE/FALSE values |
| dbByte | 2 | SMALLINT | 0-255 range |
| dbInteger | 3 | INTEGER | 32-bit integer |
| dbLong | 4 | BIGINT | 64-bit integer |
| dbCurrency | 5 | DECIMAL(19,4) | Fixed-point currency with 4 decimal places |
| dbSingle | 6 | REAL | 32-bit floating point |
| dbDouble | 7 | DOUBLE PRECISION | 64-bit floating point |
| dbDate | 8 | TIMESTAMP | Date and time (no timezone) |
| dbBinary | 9 | BYTEA | Binary data (byte array) |
| dbText | 10 | TEXT | Variable-length character string |
| dbLongBinary | 11 | BYTEA | Binary large object (byte array) |
| dbMemo | 12 | TEXT | Text large object |
| dbGUID | 15 | UUID | Universally unique identifier |
| dbBigInt | 16 | BIGINT | 64-bit integer |
| dbDecimal | 20 | DECIMAL | Decimal with configurable precision |
| dbDBTime | 21 | BIGINT | DBTime value stored as BIGINT |
| dbTime | 22 | TIME | Time only (HH:MM:SS) |
| dbTimeStamp | 23 | TIMESTAMP | Timestamp (date and time) |

## Detailed dbText (DAO Type 10) Handling

Access dbText fields in PostgreSQL are consistently mapped to TEXT:

| Scenario | PostgreSQL Type | Notes |
|----------|-----------------|-------|
| All dbText fields | TEXT | Unlimited length, fully indexable |
| Indexed dbText fields | TEXT | Supports standard indexes and full-text search |
| Required dbText fields | TEXT NOT NULL | NOT NULL constraint added |

**Important Notes:**
- PostgreSQL TEXT is unlimited in length (up to 1 GB)
- TEXT columns can be fully indexed without prefix limitations
- TEXT supports full-text search with GIN indexes
- No special handling needed for field size limits (unlike MySQL or SQL Server)

## Default Fallback

Any unrecognized DAO type code defaults to: **TEXT**

## Index Compatibility Considerations

| PostgreSQL Type | Indexable | Notes |
|-----------------|-----------|-------|
| BOOLEAN | Yes | B-tree, GIN |
| SMALLINT | Yes | B-tree |
| INTEGER | Yes | B-tree |
| BIGINT | Yes | B-tree |
| DECIMAL | Yes | B-tree |
| REAL | Yes | B-tree, GiST |
| DOUBLE PRECISION | Yes | B-tree, GiST |
| TIMESTAMP | Yes | B-tree |
| TIME | Yes | B-tree |
| TEXT | Yes | B-tree, GIN (trigram), full-text search |
| BYTEA | Yes | B-tree (equality only) |
| UUID | Yes | B-tree, hash |

## Special PostgreSQL Features

**UUID Support:**
- PostgreSQL has native UUID type
- Access GUIDs are stored as UUID
- Supports built-in UUID functions and operators

**BYTEA for Binary Data:**
- Replaces Access dbBinary and dbLongBinary
- Supports hex and escape formats
- Maximum size 1 GB

**TEXT Advantages:**
- No practical length limit
- Full-text search integration
- Pattern matching with pg_trgm extension
- Collation support per column

## Schema: public

All tables are created in the **public** schema (default PostgreSQL schema).

## Auto-increment / Identity

| Access AutoNumber | PostgreSQL Equivalent | Notes |
|-------------------|----------------------|-------|
| AutoNumber (increment) | SERIAL or BIGSERIAL | Auto-incrementing integer |
| AutoNumber (random) | UUID with DEFAULT gen_random_uuid() | For random identifiers |

**Examples:**
```sql
-- For integer AutoNumber
"PatientID" SERIAL PRIMARY KEY

-- For GUID/AutoNumber random
"RecordID" UUID DEFAULT gen_random_uuid() PRIMARY KEY