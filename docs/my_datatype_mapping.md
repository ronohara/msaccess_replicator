# MS Access to MySQL Data Type Mappings

## Summary Table

| MS Access Data Type | DAO Type Code | MySQL Data Type | Notes |
|---------------------|---------------|-----------------|-------|
| dbBoolean | 1 | BOOLEAN | Stored as TINYINT(1) in MySQL, TRUE/FALSE values |
| dbByte | 2 | TINYINT | 0-255 range |
| dbInteger | 3 | INT | 32-bit integer |
| dbLong | 4 | BIGINT | 64-bit integer |
| dbCurrency | 5 | DECIMAL(19,4) | Fixed-point currency with 4 decimal places |
| dbSingle | 6 | FLOAT | 32-bit floating point |
| dbDouble | 7 | DOUBLE | 64-bit floating point |
| dbDate | 8 | DATETIME | Date and time (no timezone) |
| dbBinary | 9 | BLOB | Binary large object, up to 65,535 bytes |
| dbText | 10 | VARCHAR(n) or TEXT | See detailed handling below |
| dbLongBinary | 11 | LONGBLOB | Binary large object, up to 4 GB |
| dbMemo | 12 | LONGTEXT | Text large object, up to 4 GB |
| dbGUID | 15 | CHAR(36) | Store as 36-character string (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx) |
| dbBigInt | 16 | BIGINT | 64-bit integer |
| dbDecimal | 20 | DECIMAL(19,4) | Decimal with 4 decimal places |
| dbDBTime | 21 | BIGINT | DBTime value stored as BIGINT |
| dbTime | 22 | TIME | Time only (HH:MM:SS) |
| dbTimeStamp | 23 | DATETIME | Timestamp (date and time) |

## Detailed dbText (DAO Type 10) Handling

Access dbText fields require special handling to choose an appropriate MySQL type:

| Field Size | MySQL Type | Notes |
|------------|------------|-------|
| 1-255 | VARCHAR(n) | Use exact field size from Access |
| 256-65535 | TEXT | Exceeds VARCHAR limit, use TEXT |
| No size defined | TEXT | Default fallback |
| Cannot read size | TEXT | Error recovery fallback |

**Important Notes:**
- VARCHAR maximum is 65,535 bytes (subject to row size limit)
- TEXT columns cannot have indexes unless prefix length is specified
- For indexed text columns, consider limiting to VARCHAR(255) or VARCHAR(450)

## Default Fallback

Any unrecognized DAO type code defaults to: **TEXT**

## Index Compatibility Considerations

| MySQL Type | Indexable | Notes |
|------------|-----------|-------|
| BOOLEAN | Yes | Fully indexable |
| TINYINT | Yes | Fully indexable |
| INT | Yes | Fully indexable |
| BIGINT | Yes | Fully indexable |
| DECIMAL | Yes | Fully indexable |
| FLOAT | Yes | Fully indexable |
| DOUBLE | Yes | Fully indexable |
| DATETIME | Yes | Fully indexable |
| TIME | Yes | Fully indexable |
| VARCHAR(n) | Yes | Index key length up to 767 bytes (InnoDB) |
| TEXT | Limited | Requires prefix length for indexing |
| LONGTEXT | Limited | Requires prefix length for indexing |
| BLOB | Limited | Requires prefix length for indexing |
| LONGBLOB | Limited | Requires prefix length for indexing |
| CHAR(36) | Yes | Fully indexable |

## Character Set

MySQL tables are created with:
- **Character Set:** `utf8mb4`
- **Collation:** `utf8mb4_unicode_ci`

This supports full Unicode, including emoji and supplementary characters.

## Storage Engine

All tables are created with:
- **Engine:** `InnoDB`

This is required for foreign key constraint support (ON DELETE CASCADE).

## Migration Notes

1. **AutoNumber in Access** → `AUTO_INCREMENT` in MySQL (must be part of PRIMARY KEY)
2. **Indexed dbText fields** with size > 255 may need manual adjustment (prefix indexes)
3. **MySQL BOOLEAN** is an alias for TINYINT(1) – TRUE=1, FALSE=0
4. **Access Yes/No fields** (dbBoolean) are converted to BOOLEAN
5. **Access Hyperlink fields** (not listed) are typically stored as dbText – use VARCHAR or TEXT
6. **Access Attachment fields** (complex) not directly supported – handle separately
7. **Access Multi-value fields** (complex) not directly supported – requires normalization

## Example CREATE TABLE Statement

```sql
CREATE TABLE `Patients` (
    `PatientID` INT AUTO_INCREMENT PRIMARY KEY,
    `Name` VARCHAR(255) NOT NULL,
    `DateOfBirth` DATETIME,
    `IsActive` BOOLEAN DEFAULT TRUE,
    `Notes` LONGTEXT,
    `Photo` LONGBLOB
) ENGINE=InnoDB CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci