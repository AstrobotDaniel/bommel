# 🐻 BOMMEL - BOM Merger for Excel Lists
 
**Author:** Daniel Leidner  
**License:** MIT License (commercial use allowed)  
**Version:** 1.0
 
---
 
## 📋 What is BOMMEL?
 
BOMMEL is a smart Excel list merger designed specifically for Bill of Materials (BOM) and procurement lists. It automatically transfers manually maintained data (like procurement status, supplier info, prices) from your old list into a new updated list, even when the structure changes.
 
**The Problem:** You have an old BOM with manually added procurement data (status, prices, suppliers, notes). A new version arrives with updated parts, renamed files, or structural changes. Manually copying everything over is tedious and error-prone.
 
**The Solution:** BOMMEL intelligently matches entries between old and new lists and transfers your manual work automatically.
 
---
 
## ✨ Key Features
 
### 🎯 Intelligent Matching
- **Configurable matching key** set at top of script
- **Exact key comparison** — rows with identical keys are matched across versions
- **Reports match statistics** (matched, updated, new, deleted items)
 
### 📊 Smart Data Transfer
- **Transfers manual columns:** using the informtion provided by the user
- **Inserts missing columns** using "left neighbor" logic (finds where to place them)
- **Preserves formatting** for assembly header rows 
- **Preserves formulas** from the new workbook with correct column adjustment
- **Sets new entries** to Status="new" for manual review
 
### 🎨 Conditional Formatting (Auto-updating Colors)
- **⚪ WHITE** = `-` (no status)
- **🟠 CORAL** = `new` (manual review needed)
- **🟡 YELLOW** = `requested` (inquiry sent)
- **🟣 PURPLE** = `offered` (offer received)
- **🔵 LIGHT BLUE** = `ordered`
- **🟢 LIGHT GREEN** = `paid`
- **🟢 MEDIUM GREEN** = `delivered`
- **🟢 DARK GREEN** = `completed` + Bold
- **⚫ GRAY** = `postponed` + Italic
- **🔴 LIGHT RED** = quantity changed on an already-ordered item (highest priority)
 
Colors update automatically when you change the Status dropdown!
 
### 📁 Output Structure
**Main Sheet ("Main_List"):**
- Merged list with all data
- Status dropdown in each row
- AutoFilter enabled
- Auto-width columns
- Tracking column: `_status_flag` (NEW / UPDATED)
 
**Deleted Items Sheet ("Deleted_Entries"):**
- Items from old list not in new list
- Red background for visibility
 
**Log Sheet:**
- Timestamp
- File names
- Match statistics
- Column transfer summary
 
---
 
## 🚀 Usage
 
**Setup:**
1. Open **both** Excel files (old + new)
2. Press `Alt+F11` to open VBA Editor
3. Go to `Insert` → `Module`
4. Paste BOMMEL code into module
5. Update filenames in lines 30-31:
   ```vba
   Const OLD_WORKBOOK_NAME As String = "POD.02_BOM_mech_vorlaeufig_20260327.xlsx"
   Const NEW_WORKBOOK_NAME As String = "POD.02_20260413.xlsx"
   ```
6. Press `F5` or click `Run`
 
**Features:**
- Works entirely in Excel Desktop
- No external dependencies
- Debug output in Immediate Window (`Ctrl+G`)
- Creates new workbook with results
 
 
---
 
## 🛡️ Safety Features
 
- **New workbook output** (never modifies original files)
- **Deleted items preserved** in a separate sheet
- **Tracking columns** show what was changed (`_status_flag`: NEW / UPDATED)
- **First sheet used** by name-agnostic detection (works regardless of sheet name)
 
 