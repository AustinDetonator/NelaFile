# NelaFile Developer Guide

## 1. Architecture Overview

### Technology Stack
- **Framework**: .NET Framework 4.8 (WinForms)
- **Database**: 
  - Oracle (IFS): `System.Data.OracleClient` + `Oracle.ManagedDataAccess.Client`
  - SQL Server (Audit): `System.Data.SqlClient`
- **XML**: `System.Xml` (XmlDocument, XPath)
- **File I/O**: `System.IO` (StreamWriter, file copy)
- **Network**: FTP client for IMT export

### Key Classes
- `Form1`: Main UI and business logic
- `Funkce`: Oracle helper methods
- `Database`: SQL Server utilities
- `FtpClient`: IMT FTP operations
- `Settings` / `Globals`: Configuration

---

## 2. Database Schema (Key Tables)

### Oracle IFS (AD9PROD / AD9TEST)
```
IFSAPP.C_PACKAGE_HEAD         -- order package header
IFSAPP.C_PACKAGE_DETAIL       -- order package detail lines
IFSAPP.A0000_CustomerByOrderNo -- customer/order lookup
IFSAPP.INV_PART_DISCRETE_CHAR -- product characteristics (nominal time, delay, etc.)
IFSAPP.c_labels_ad            -- label files (KTN/STN/PRO)
IFSAPP.EDM_FILE_STORAGE       -- BLOB storage for XML documents (doc_class='ZO-XML')
IFSAPP.A1135_INVPART_CHAR_ALL_CFV -- part characteristics (UFI codes)
IFSAPP.DOC_REFERENCE_OBJECT   -- document references
IFSAPP.DOC_STRUCTURE          -- document structure
IFSAPP.DOC_ISSUE_REFERENCE    -- document issues/versions
```

### SQL Server (AUSTIN_SLC)
```
nela                          -- export log table
NewAI                         -- label-to-application-ID mapping
```

---

## 3. Core Method Flow

### VytvorSoubor() – Main NELA/Standard Flow
Located: `Form1.cs:503`

Flow:
1. Open Oracle connection to IFS
2. Check if "economic packaging" (`unique_pkg='FALSE'`)
3. Query country for order
4. If Ukraine/Alter CK enabled:
   - Query alternate barcode mapping
   - Resolve alternative order number
5. Query and collect label files (KTN/STN/PRO)
6. Call stored procedure `IFSAPP.A0055_ONLINENELA.PRONELU` to get order data
7. For each data row:
   - Build TXT lines with order/package/catalog info
   - **If Alter CK active**: call `XML_ExLabeling_Get_Part_no()` to resolve alternate part number from XML
   - Write all lines to StreamWriter
8. Log export metadata to SQL Server
9. Call `KopirujNalepky()` to copy label files

### XML_ExLabeling_Get_Part_no() – XML Resolution
Located: `Form1.cs:5547`

**Root Cause of Exception**: Fetches BLOB from `IFSAPP.EDM_FILE_STORAGE` using broad `LIKE` query; may return non-XML or wrong document; `LoadXml()` fails if content is malformed or not UTF-8.

**Fix Strategy**:
- Use exact document lookup (not `LIKE`)
- Stream-load XML with auto-detect encoding
- Parameterize SQL

```csharp
private string XML_ExLabeling_Get_Part_no(string pOrder_no, string pPart_no)
{
    string sPart_no = string.Empty;

    // Use parameterized query for exact doc lookup
    string sSQL_Get_XML_BLOB =
       "SELECT file_data FROM IFSAPP.EDM_FILE_STORAGE a " +
       "WHERE a.doc_class = 'ZO-XML' " +
       "AND a.doc_no = :docNo " +
       "AND ROWNUM = 1";  // deterministic

    using (var conn = new Oracle.ManagedDataAccess.Client.OracleConnection(Globals.ConString))
    {
        conn.Open();

        using (var cmd = new Oracle.ManagedDataAccess.Client.OracleCommand(sSQL_Get_XML_BLOB, conn))
        {
            // Parameterize to prevent SQL injection + ensure exact match
            cmd.Parameters.Add(":docNo", OracleDbType.Varchar2).Value = pOrder_no;

            using (var reader = cmd.ExecuteReader())
            {
                if (reader.Read() && reader["file_data"] != DBNull.Value)
                {
                    byte[] blobBytes = (byte[])reader["file_data"];

                    // Stream-load with auto-detect encoding (solves UTF-8 BOM issues)
                    var xmlDoc = new XmlDocument();
                    using (var ms = new MemoryStream(blobBytes))
                    {
                        ms.Position = 0;
                        xmlDoc.Load(ms);  // XmlDocument.Load(Stream) auto-detects encoding
                    }

                    sPart_no = XML_GetElement_SID(xmlDoc, pPart_no);
                }
            }
        }
    }

    return sPart_no;
}
```

---

## 4. Exception: System.Xml.XmlException

### Root Cause
- BLOB content is not valid XML (binary data, wrong encoding, corrupted)
- UTF-8 hardcoded + BOM mismatch
- Wrong document retrieved (broad `LIKE` query returns non-XML row)

### Stack Trace
```
System.Xml.XmlException: Data at root level is invalid.
  at System.Xml.XmlTextReaderImpl.ParseRootLevelWhitespace()
  at System.Xml.XmlTextReaderImpl.ParseDocumentContent()
  at System.Xml.XmlLoader.Load(XmlDocument, XmlReader, bool)
  at System.Xml.XmlDocument.LoadXml(string)
  at NelaFile.Form1.XML_ExLabeling_Get_Part_no(...) [line 5578]
```

### Fix Applied
1. Use `XmlDocument.Load(Stream)` instead of `LoadXml(string)` → auto-detect encoding
2. Parameterize SQL to ensure exact document match
3. Add null/DBNull checks for blob content

---

## 5. Configuration

### Globals (assumed class/static settings)
```csharp
public static string ConString         // Oracle prod/test connection string
public static string ConString_IFS9    // Legacy IFS v9 connection string
public static string ConString_AUSTIN_SLC  // SQL Server audit DB
public static string TNSNAME           // Oracle TNS name
public static IFS_Versions IFS_VERSION // Current IFS version (IFS_9 or other)
public static string FOLDER_E3PR_DEFAULT      // Output folder for E3PR
public static string FOLDER_DA001_DEFAULT     // Output folder for DA001
public static string FOLDER_NEW_VERSION       // Version check path
```

### Connection Strings (from app.config)
Example:
```xml
<add name="AD9PROD" value="Data Source=AD9PROD;User Id=lotyl;Password=...;"/>
<add name="AD9TEST" value="Data Source=AD9TEST;User Id=lotyl;Password=...;"/>
```

---

## 6. Known Issues & Recommendations

### Issue 1: SQL Injection Risk
**Problem**: All queries use string concatenation:
```csharp
"... WHERE a.order_no = '" + _Customer_order + "'"
```

**Recommendation**: Migrate to parameterized queries (OracleParameter, SqlParameter).

### Issue 2: Mixed Oracle Providers
**Problem**: App uses both `System.Data.OracleClient` (legacy) and `Oracle.ManagedDataAccess`.

**Recommendation**: Consolidate to `Oracle.ManagedDataAccess` only (supported, faster).

### Issue 3: Tight UI Coupling
**Problem**: Business logic, DB calls, file I/O directly in Form1_Click handlers.

**Recommendation**: Extract to service layer (e.g., `NelaService`, `LabelService`).

### Issue 4: No Async/Threading
**Problem**: Long DB queries block UI thread.

**Recommendation**: Use `async/await` with `Task` for DB operations.

### Issue 5: Exception Handling
**Problem**: Broad catch blocks hide root causes; XML errors silently return empty string.

**Recommendation**: Log exception details; re-throw or handle specifically.

---

## 7. Testing Checklist

### Unit Tests (recommended)
- `XML_ExLabeling_Get_Part_no()` with valid/invalid/empty BLOB
- `XML_GetElement_SID()` with XPath variations
- `VytvorSoubor()` with mock Oracle/SQL connections

### Integration Tests
- End-to-end: order load → TXT generation → file copy
- Alter CK flow with Ukraine order
- FTP copy functionality

### Manual QA
- [ ] Standard order (no Alter CK)
- [ ] Ukraine order (Alter CK enabled)
- [ ] Missing label scenarios
- [ ] Invalid XML document scenarios
- [ ] Network/DB connectivity loss

---

## 8. Building & Debugging

### Prerequisites
```
Visual Studio 2019+ with:
  - .NET Framework 4.8 SDK
  - Oracle.ManagedDataAccess NuGet
  - System.Data.OracleClient (legacy)
```

### Build
```bash
dotnet build
# or in Visual Studio: Build → Build Solution
```

### Debug
1. Set breakpoint in `XML_ExLabeling_Get_Part_no()` at line 5578
2. Run with debugger (F5)
3. Inspect `xmlString` value and `blobBytes` hex
4. Check Oracle SQL Profiler for query results

### Common Debug Points
- `VytvorSoubor()` line 986: XML method entry
- `XML_ExLabeling_Get_Part_no()` line 5574: encoding conversion
- `XML_GetElement_SID()` line 5579: XPath evaluation

---

## 9. Deployment

### Environment Setup
1. Target SQL Server with `nela` table
2. Target Oracle instance with IFS app schema access
3. Configure `app.config` with correct connection strings
4. Create output folders with write permissions

### Release Build
```bash
dotnet publish -c Release -f net48
```

Copy to target machine; run `NelaFile.exe`.

---

## 10. Support & Contacts
- **Code owner**: [TBD]
- **Oracle DBA**: [TBD]
- **SQL Server DBA**: [TBD]
- **FTP admin**: [TBD]
