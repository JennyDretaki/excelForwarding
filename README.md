# Excel Importer — Application Report

## Purpose

A C# Windows Forms application that imports data from Excel files (`.xls` / `.xlsx`) and creates **shortcuts (`.lnk`)** in destination folders, based on matching fields from each row.

---

## Workflow

1. **Start screen:** waiting panel for drag & drop (or the “Add Excel...” button).
2. **Drag visual feedback:**
   - green panel → valid file (`.xls` / `.xlsx`)
   - red panel → invalid file
3. **Import:** reads headers and rows from Excel.
4. **Preview:** shows results in a grid — **no** send yet.
5. **Send:** shortcuts are created in the destination folders only after user confirmation.

---

## Columns shown in Preview

| Column | Description |
|--------|-------------|
| Court | From header `ΔΙΚΑΣΤΗΡΙΟ` |
| Tax ID (AFM) | From header `ΑΦΜ` |
| Host Account | From header `Host Account` |
| Comments | From header `ΣΧΟΛΙΟ/ΕΝΕΡΓΕΙΑ` |
| GAK | From header `ΓΑΚ` |
| EAK | From header `ΕΑΚ` |
| Date | From header `ΗΜΕΡΟΜΗΝΙΑ ΔΙΚΑΣΙΜΟΥ` |
| Message | Status / error per row |

The Excel file must also include **Full name** (`ΟΝΟΜΑΤΕΠΩΝΥΜΟ`) for matching — it is not shown in the grid.

---

## Excel import rules

- Only `.xls` and `.xlsx` are accepted.
- If required headers are missing, a message is shown:  
  **“Not all fields were found (…)”**.
- Drag & drop is rejected if valid and invalid files are mixed.

---

## Destination folder matching

For each row:

1. Lookup by **Host Account** (contract / live contract equivalent).
2. Confirm with **GAK** and the **first 5 characters** of the full name.
3. Validate folder consistency across the two data sources.
4. Shortcut destination:

```text
{ScannedRoot}\{newfakelos2}\{fakelos}\
```

`ScannedRoot` is defined in application settings; `newfakelos2` and `fakelos` come from the matching process.

- **Folders are not created** — only existing paths are used.
- If the shortcut already exists, **it is not created again**.

---

## Sending shortcuts

- A `.lnk` file is created locally, pointing to the original Excel file.
- It is written to the destination folder using secured network access (impersonation), consistent with related internal tools.
- Network paths use UNC form (not a mapped drive letter) so access works correctly under impersonation.

The **Send** button is enabled only when at least one row is in **Ready** status.

---

## Preview statuses

| Status | Color | Meaning |
|--------|--------|---------|
| Ready | Green | Can be sent |
| Already exists / Duplicate | Yellow | Skipped |
| Error | Orange | Not sent — details in the Message column |

---

## Configuration (`appsettings.json`)

- **Connection string** to the business folder database.
- **ScannedRoot:** root of scanned folders (UNC form preferred).
- **Impersonation:** domain / user / password for network write access.

---

## Technologies

- .NET 8 Windows Forms  
- ClosedXML / ExcelDataReader (Excel reading)  
- Microsoft.Data.SqlClient  
- Microsoft.Extensions.Configuration.Json  

---

## How to run

```powershell
cd ExcelImporter
dotnet run
```

Or open `ExcelImporter.csproj` and run from the IDE.

---

## Summary

The application provides controlled Excel import, preview of results, and targeted placement of shortcuts into existing folders, with validation and without duplicate shortcuts.
