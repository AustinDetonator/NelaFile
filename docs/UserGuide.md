# NelaFile User Guide

## 1. Purpose
NelaFile is used to generate TXT files for production machines (NELA, E3PR, DA001), collect related label files, and export them to the target folder (and optionally FTP).

---

## 2. Requirements
Before use, confirm:
- Network access to IFS resources and shared folders
- Database connectivity (Oracle and SQL Server)
- Correct application environment selected (production/test)
- Valid output folder configured

---

## 3. Main Screen Overview
Key controls:
- **Load orders**: fills the order grid
- **Order grid (`dataGridViewVZ`)**: source order selection
- **Machine selector (`comboBoxNelaCislo`)**: target machine
- **Output folder**: destination for generated TXT and copied labels
- **Create file** button: generates TXT output
- **Labels grid (`dataGridViewNalepky`)**: related labels for selected order
- **Files grid (`dataGridViewSoubory`)**: export history

---

## 4. Standard Workflow

### Step 1: Load orders
1. Start NelaFile.
2. Click **Load orders**.
3. Wait until the order list is filled.

### Step 2: Select order
1. Select one row in the order grid.
2. Confirm the order data is correct.

### Step 3: Select machine
1. In **Machine selector**, choose target machine:
   - Nela1 / Nela2 / Nela3 / Nela4
   - E3PR
   - DA001

### Step 4: Set output folder
1. Click folder browse button.
2. Select a valid writable folder.
3. Confirm path appears in output folder textbox.

### Step 5: Generate file
1. Click **Create file**.
2. Wait for completion message:
   - `Soubor je vytvořen.`
3. The app writes TXT and copies label files.

### Step 6: Verify output
1. Open target folder.
2. Confirm:
   - TXT file exists
   - `Export` subfolder contains label files

---

## 5. Optional Operations

### Open generated TXT from history
1. Select a row in files history grid.
2. Use context menu **Open TXT**.

### Edit generated TXT
1. Select a row in files history grid.
2. Use context menu **Edit TXT**.

### Copy to IMT FTP (when available)
1. Select exported row for target machine.
2. Use IMT FTP action from context menu.
3. Confirm successful copy message.

---

## 6. Common Options
- **Only stock orders**: changes order source/validation behavior
- **Decimal delimiter change**: forces dot/comma formatting in output values
- **Alter CK / UA flow**: for alternative barcode/order behavior when applicable

Use these options only when required by process instruction.

---

## 7. Troubleshooting

### "Vyberte řádek"
No order row selected. Select one order row and retry.

### "Není zadána NELA nebo stroj."
No machine selected. Choose machine and retry.

### Label not found errors (`Nenašel nálepky ...`)
Required labels are missing in IFS/DMS for selected order line.
- Check order attachments
- Check doc class availability (KTN/STN/PRO)

### XML parse error (root-level invalid data)
The alternate XML document content is invalid or wrong document was resolved.
- Verify correct XML is attached for the order
- Verify XML storage content and encoding
- Contact support if persistent

### File/folder access problems
Ensure user has write access to output and export directories.

---

## 8. Best Practices
- Use production/test environment carefully.
- Verify selected order and machine before generating.
- Keep output folder on reliable network path.
- Review generated TXT before sending to production.

---

## 9. Support Information (fill in internally)
- Application owner:
- Production support:
- DBA contact:
- FTP support:
