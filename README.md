# Circuit Diagram Inspector

Interactive PDF inspection tool for circuit diagrams with drag-to-annotate functionality and automatic error logging to Excel.

## Features

✅ **Interactive PDF Viewer**
- Load and view circuit diagram PDFs
- Zoom in/out (0.5x - 3.0x) for detailed inspection
- Multi-page navigation with smooth scrolling
- High-resolution rendering (2x) for clarity

✅ **Drag-to-Annotate**
- **Left Click & Drag**: Draw green circles to mark components as OK
- **Right Click & Drag**: Draw yellow rectangles to highlight errors
- Real-time preview while dragging
- Adjustable sizes for both circles and rectangles

✅ **Annotation Management**
- **Ctrl+Click**: Select any annotation (highlighted with blue border)
- **Delete Key**: Remove selected annotation with confirmation
- Visual feedback for selected annotations
- Export annotated PDF with all markings visible

✅ **Hierarchical Error Categories**
- **Material Shortfall**: Components missing, terminals missing, short link missing
- **Wrong Wiring**: Wire interchanges, color code wrong, ferrule wrong, size wrong, loose wires, lug issues
- **Incomplete Wiring**: Incomplete connections, pending wiring
- **Wrong Assembly**: Wrong labels, fuses, wire ducts, components, fixation issues, loose end stoppers

✅ **Excel Integration (Emerson.xlsx)**
- Automatic logging with sequential Sr No (continues across sessions)
- Merged cell headers: Project Name, Sales Order, Cabinet ID
- Detailed punch list with category, checked by name, and date
- Auto-formatted cells with borders and alignment
- Tag-based error descriptions (e.g., "TB1 missing", "R5 Fuse Wrong installed")

✅ **PDF Export**
- Export annotated PDFs with visible green circles and yellow rectangles
- Save location of your choice
- Annotations burned into PDF permanently

## Installation

### 1. Install Python
Make sure you have Python 3.8 or higher installed.

### 2. Clone the Repository

```powershell
git clone https://github.com/Veda2254/Circuit-diagram-inspector.git
cd Circuit-diagram-inspector
```

### 3. Create Virtual Environment (Recommended)

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

### 4. Install Dependencies

```powershell
pip install -r requirements.txt
```

Or install individually:
```powershell
pip install PyMuPDF Pillow openpyxl
```

## Usage

### Running the Application

```powershell
python circuit_inspector.py
```

### Workflow

1. **Load PDF**: Click "📁 Load PDF" to select your circuit diagram
   - PDF loads immediately without prompts
   - Project details (name, sales order, cabinet ID) can be entered manually in Excel

2. **Set Cabinet ID** (Optional): Click "🆔 Set Cabinet ID" to update Excel headers
   - Enter Project Name, Sales Order Number, and Cabinet ID
   - This updates the merged cells in Emerson.xlsx

3. **Inspect Components**:
   
   **Mark OK Components:**
   - Hold **Left Mouse Button** and drag to create a green circle
   - Circle size adjusts based on drag distance from center
   - Release to finalize
   
   **Mark Errors:**
   - Hold **Right Mouse Button** and drag to create a yellow rectangle
   - Rectangle size adjusts based on drag area
   - Release to show tag input dialog
   - Enter tag details (e.g., "TB1", "R5", "F1")
   - Select error category from hierarchical menu
   - Error automatically logged to Excel

4. **Manage Annotations**:
   - **Ctrl+Click** on any annotation to select it (blue border appears)
   - Press **Delete** key to remove selected annotation
   - Annotations stay on screen for reference

5. **Navigate**: Use "◀ Prev" and "Next ▶" buttons to move between pages

6. **Zoom**: Use "🔍+" and "🔍-" for detailed inspection (0.5x to 3.0x)

7. **Export**: Click "📥 Export Annotated PDF" to save PDF with all visible annotations

8. **View Log**: Click "📊 Open Excel" to see the complete error log

### Excel Output Format (Emerson.xlsx)

**Headers:**
- E4:M4 - Project Name
- E6:H6 - Sales Order No
- K6:L6 - Cabinet ID

**Data Table (Starting Row 11):**
| Sr No | Reference No | Punch / Action Point | Category | Checked By (Name) | Checked By (Date) | Implemented By | Closed By |
|-------|--------------|----------------------|----------|-------------------|-------------------|----------------|-----------|
| 1 | | TB1 missing | Material Shortfall | Username | 2025-11-25 | | |
| 2 | | R5 Fuse Wrong installed | Wrong Assembly | Username | 2025-11-25 | | |
| 3 | | Wire X3-5 Color Code Wrong | Wrong Wiring | Username | 2025-11-25 | | |

**Features:**
- Sr No auto-increments sequentially across sessions
- Punch text auto-generated from tag and error type
- Checked By name and date filled automatically
- All cells properly formatted with borders and alignment

## Keyboard Shortcuts

| Key | Action |
|-----|--------|
| **Ctrl+Click** | Select annotation |
| **Delete** | Remove selected annotation |

## Error Category Structure

The tool includes comprehensive error categories:

### Material Shortfall
- Missing components (e.g., "TB1 missing")
- Terminals missing (e.g., "X3 Terminal Missing")
- Short link missing (e.g., "TB5 TB Group 1-2 shortlink missing")

### Wrong Wiring
- Wires Interchanges
- Color Code Wrong
- Ferrule Wrong
- Size Wrong
- Wire Loose Found
- Lug not properly cut

### Incomplete Wiring
- All wiring Incomplete with connections pending
- Connections pending

### Wrong Assembly
- Label Wrong installed
- Fuse Wrong installed
- Wire duct Wrong Installed
- Component Wrong installed
- Component not properly fixed
- End stopper loose found

## Customization

### Adding New Error Types

Edit the `error_categories` dictionary in `circuit_inspector.py`:

```python
self.error_categories = {
    'Material Shortfall': {
        'Missing': '{tag} missing',
        'Your New Error': '{tag} your custom text'
    },
    'Your Category': {
        'Error Type': '{tag} error description'
    }
}
```

Use `{tag}` placeholder for component name and `{terminals}` for terminal numbers (short link errors).

### Changing Excel File

Modify the `excel_file` variable in the `__init__` method:

```python
self.excel_file = "YourTemplate.xlsx"
```

Ensure your Excel file has the same structure as Emerson.xlsx.

## Troubleshooting

**Issue**: PDF doesn't load
- Make sure the PDF file is not corrupted
- Check if you have read permissions for the file
- Verify the file is a valid PDF

**Issue**: Excel file access error
- Close the Excel file if it's open in another program
- Check write permissions for the directory
- Ensure Emerson.xlsx exists in the same folder

**Issue**: Merged cells not updating
- Close Excel completely before updating headers
- The app automatically unmerges, writes, and re-merges cells

**Issue**: Annotations not visible
- Check you're on the correct page
- Selected annotations show with blue border
- Green circles = OK, Yellow rectangles = Errors

**Issue**: Sr No not sequential
- The app reads the last Sr No from Excel automatically
- If Excel is corrupted, delete and recreate Emerson.xlsx

**Issue**: Blurry PDF display
- Use the zoom controls (🔍+ / 🔍-) to adjust view
- The app renders at 2x resolution for clarity

## Tips

✨ **Workflow Tips**
- You can start annotating immediately after loading PDF
- Cabinet ID can be set anytime (updates Excel headers)
- Drag size determines circle radius and rectangle size
- Use Ctrl+Click to review and correct annotations
- Export annotated PDF for sharing with team

✨ **Excel Tips**
- Emerson.xlsx is created automatically on first run
- Sr No continues from last entry (never resets to 1)
- Fill Project Name, Sales Order, Cabinet ID manually in Excel if preferred
- Excel remains open - close and reopen to see updates

✨ **Efficiency Tips**
- Use left drag for quick OK marking (green circles)
- Right drag for detailed error reporting (yellow boxes)
- Zoom in for small components
- Navigate pages with toolbar buttons
- Delete key for quick annotation removal

## System Requirements

- **OS**: Windows 10/11 (PowerShell)
- **Python**: 3.8 or higher (tested with 3.11.9)
- **RAM**: 4GB minimum, 8GB recommended
- **Display**: 1280x720 or higher recommended
- **Storage**: 100MB for dependencies

## Dependencies

```
PyMuPDF==1.23.8      # PDF rendering and manipulation
Pillow==10.1.0        # Image processing and annotation drawing
openpyxl==3.1.2       # Excel file handling
opencv-python==4.8.1.78  # Image processing utilities
numpy==1.24.3         # Array operations
tkinter               # GUI framework (usually included with Python)
```

## Project Structure

```
Circuit-diagram-inspector/
├── circuit_inspector.py    # Main application (810 lines)
├── Emerson.xlsx            # Excel template for logging
├── requirements.txt        # Python dependencies
├── README.md              # This file
├── .gitignore            # Git ignore patterns
└── .venv/                # Virtual environment (not in repo)
```

## Version History

**Current Version: 2.0**
- ✅ Drag-to-draw green circles and yellow rectangles
- ✅ Annotation selection and deletion (Ctrl+Click + Delete)
- ✅ Sequential Sr No across sessions
- ✅ PDF export with visible annotations
- ✅ Hierarchical error categories
- ✅ Tag-based punch text generation
- ✅ No project dialogs on load (manual Excel entry)
- ✅ Fixed merged cell Excel updates

**Previous Version: 1.0**
- Basic click annotations
- Fixed-size highlights
- JSON annotation storage

## Contributing

For bug reports or feature requests, please contact the development team or create an issue in the repository.

## License

Free to use for internal production and quality control purposes.
