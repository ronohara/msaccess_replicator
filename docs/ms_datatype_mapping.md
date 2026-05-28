# MS Access to SQL Server Data Type Mappings

## Summary Table

| MS Access Data Type | DAO Type Code | SQL Server Data Type | Notes |
|---------------------|---------------|----------------------|-------|
| dbBoolean | 1 | BIT | 0 or 1 (True/False) |
| dbByte | 2 | TINYINT | 0-255 range |
| dbInteger | 3 | INT | 32-bit integer |
| dbLong | 4 | BIGINT | 64-bit integer |
| dbCurrency | 5 | DECIMAL(19,4) | Fixed-point currency with 4 decimal places |
| dbSingle | 6 | FLOAT | 32-bit floating point |
| dbDouble | 7 | FLOAT | 64-bit floating point |
| dbDate | 8 | DATETIME | Date and time |
| dbBinary | 9 | VARBINARY(MAX) | Binary data, up to 2 GB |
| dbText | 10 | NVARCHAR(n) or NVARCHAR(MAX) | See detailed handling below |
| dbLongBinary | 11 | VARBINARY(MAX) | Binary large object, up to 2 GB |
| dbMemo | 12 | NVARCHAR(MAX) | Text large object, up to 2 GB |
| dbGUID | 15 | UNIQUEIDENTIFIER | Globally unique identifier |
| dbBigInt | 16 | BIGINT | 64-bit integer |
| dbDecimal | 20 | DECIMAL(19,4) | Decimal with 4 decimal places |
| dbDBTime | 21 | BIGINT | DBTime value stored as BIGINT |
| dbTime | 22 | TIME | Time only (HH:MM:SS) |
| dbTimeStamp | 23 | DATETIME | Timestamp (date and time) |

## Detailed dbText (DAO Type 10) Handling

Access dbText fields require special handling to choose an appropriate SQL Server type based on field size:

| Field Size | SQL Server Type | Notes |
|------------|-----------------|-------|
| 1-4000 | NVARCHAR(n) | Use exact field size from Access, can be indexed |
| No size defined | NVARCHAR(255) | Safe default for indexing |
| Cannot read size | NVARCHAR(255) | Error recovery fallback |
| Size > 4000 | NVARCHAR(255) | Limited to 255 for index compatibility |

Important Notes:
- SQL Server NVARCHAR uses 2 bytes per character (Unicode)
- Maximum row size is 8,060 bytes for indexed columns
- NVARCHAR(MAX) cannot be used in indexes
- For indexed text columns, limit to NVARCHAR(450) to stay within 900-byte index key limit
- VARCHAR is not used to preserve Unicode support (Access uses UTF-16)

## Default Fallback

Any unrecognized DAO type code defaults to NVARCHAR(255)

## Index Compatibility Considerations

| SQL Server Type | Indexable | Notes |
|-----------------|-----------|-------|
| BIT | Yes | Fully indexable |
| TINYINT | Yes | Fully indexable |
| INT | Yes | Fully indexable |
| BIGINT | Yes | Fully indexable |
| DECIMAL | Yes | Fully indexable |
| FLOAT | Yes | Fully indexable |
| DATETIME | Yes | Fully indexable |
| TIME | Yes | Fully indexable |
| UNIQUEIDENTIFIER | Yes | Fully indexable |
| NVARCHAR(n) | Yes | Subject to 900-byte key length limit |
| NVARCHAR(MAX) | No | Cannot be used as index key |
| VARBINARY(MAX) | No | Cannot be used as index key |

## Filtered Unique Indexes for Nullable Columns

SQL Server does not allow multiple NULL values in unique indexes (unlike MS Access). To match Access behavior, filtered unique indexes are automatically created when all columns in a unique index are nullable. A filtered unique index uses a WHERE clause to exclude NULL values, allowing multiple rows with NULL while still enforcing uniqueness on non-NULL values.

## Schema

All tables are created in the dbo schema (default SQL Server schema)

## Identity/Auto-increment

Access AutoNumber (increment) converts to IDENTITY(1,1) in SQL Server

## Migration Notes

1. AutoNumber in Access converts to IDENTITY(1,1) in SQL Server
2. SQL Server BIT accepts 0 or 1 (Access Yes/No converts correctly)
3. Access Yes/No fields (dbBoolean) convert to BIT
4. Access Hyperlink fields (not listed) are typically stored as dbText – use NVARCHAR
5. Access Attachment fields (complex) are not directly supported – handle separately
6. Access Multi-value fields (complex) are not directly supported – requires normalization
7. NVARCHAR preserves Unicode (Access uses UTF-16) while VARCHAR does not
8. Indexed text columns must use NVARCHAR(n) not NVARCHAR(MAX) for successful index creation
9. Text columns that need full Unicode support should use NVARCHAR

## MERGE / UPSERT Support

SQL Server uses MERGE statement for UPSERT operations. When a primary key or unique constraint exists, the MERGE statement handles both INSERT for new rows and UPDATE for existing rows in a single atomic operation. The ON clause specifies how to match source and target rows, typically using the primary key columns.

## ON DELETE CASCADE

Foreign keys created by the replicator include ON DELETE CASCADE to maintain referential integrity. When a parent row is deleted, all related child rows are automatically deleted. This matches the behavior of MS Access relationships and ensures data consistency during deletion synchronization.

## Example Type Conversion

An Access table with the following columns:
- A Yes/No field (dbBoolean) becomes BIT
- A text field of size 50 (dbText) becomes NVARCHAR(50)
- A long text memo field (dbMemo) becomes NVARCHAR(MAX)
- An AutoNumber (dbLong with AutoNumber attribute) becomes INT IDENTITY(1,1)
- A date/time field (dbDate) becomes DATETIME