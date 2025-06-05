import xlwings as xw
import os

# === Configuration ===
xl_file = "budgeting-workbook/Budget Template - MASTER UNLOCKED (TEST).xlsm"
output_file = "budgeting-workbook/Budget Template - FINAL TEST.xlsm"

# Full module paths
modules = [
    "budgeting-workbook/vba/ResetModule.bas",
    "budgeting-workbook/vba/UpdateLevelsModule.bas",
    "budgeting-workbook/vba/SummaryFormulas.bas",
]
thisworkbook_cls = "budgeting-workbook/vba/ThisWorkbook.cls"

# === Main Logic ===
app = xw.App(visible=False)
wb = app.books.open(xl_file)

# Clear old modules
wb.api.VBProject.VBComponents.Remove(wb.api.VBProject.VBComponents("Module1"))
wb.api.VBProject.VBComponents.Remove(wb.api.VBProject.VBComponents("Module2"))

# Import new modules
for mod in modules:
    wb.api.VBProject.VBComponents.Import(os.path.abspath(mod))

# Replace ThisWorkbook
wb.api.VBProject.VBComponents.Remove(wb.api.VBProject.VBComponents("ThisWorkbook"))
wb.api.VBProject.VBComponents.Import(os.path.abspath(thisworkbook_cls))

# Save new version
wb.save(os.path.abspath(output_file))
wb.close()
app.quit()

print(f"✅ Done. Saved as: {output_file}")
