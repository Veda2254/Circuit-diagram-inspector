# Circuit Diagram Inspector

Interactive PDF inspection tool for circuit diagrams with error logging to Excel.

## Features

✅ **Interactive PDF Viewer**
- Load and view circuit diagram PDFs
- Zoom in/out for detailed inspection
- Multi-page navigation

✅ **Error Tracking**
- **Left Click**: Mark component as OK (green checkmark)
- **Right Click**: Report error with context menu
- Categorized error types for different components

✅ **Error Categories**
- **Wire**: Wire wrong, Ferrule direction wrong, Wiring not present
- **Fuse**: Fuse missing, Wrong fuse rating, Fuse orientation wrong
- **Component**: Missing component, Wrong material installation, Missing material, Wrong component type
- **General**: Assembly error, Labeling error, Connection loose, Other

✅ **Excel Logging**
- Automatic logging to `inspection_log.xlsx`
- Tracks: Timestamp, Cabinet ID, Page, Component Type, Error Description, Inspector

✅ **Annotation Persistence**
- Save annotations to JSON file
- Visual markers on PDF (green for OK, yellow for errors)

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
2. **Set Cabinet ID**: Click "🆔 Set Cabinet ID" to enter the cabinet identifier (e.g., X-11-1)
3. **Inspect Components**:
   - **Left Click** on components that are correct → Green checkmark appears
   - **Right Click** on errors → Enter component name (e.g., "F1 fuse", "Wire X3-5") → Select error type from menu → Yellow marker appears
4. **Navigate**: Use "◀ Prev" and "Next ▶" buttons to move between pages
5. **Zoom**: Use "🔍+" and "🔍-" for detailed inspection
6. **Save Work**: Click "💾 Save Annotations" to save your progress
7. **View Log**: Click "📊 Open Excel" to see the error log

### Excel Output Format

The error log includes the component name in the error description:

| Timestamp | Cabinet ID | Page | Component Type | Error Description (Component + Error) | Inspector |
|-----------|------------|------|----------------|---------------------------------------|-----------|
| 2025-11-24 10:30:45 | X-11-1 | 1 | Wire | xyz fuse Wire wrong | YourUsername |
| 2025-11-24 10:31:12 | X-11-1 | 2 | Fuse | F1 fuse Fuse missing | YourUsername |
| 2025-11-24 10:31:12 | X-11-1 | 2 | Fuse | Fuse missing | YourUsername |

## Customization

### Adding New Error Types

Edit the `error_categories` dictionary in `circuit_inspector.py`:

```python
self.error_categories = {
    'Wire': ['Wire wrong', 'Ferrule direction wrong', 'Wiring not present'],
    'Fuse': ['Fuse missing', 'Wrong fuse rating', 'Fuse orientation wrong'],
    'YourCategory': ['Error 1', 'Error 2', 'Error 3'],
}
```

### Changing Excel File Location

Modify the `excel_file` variable in the `__init__` method:

```python
self.excel_file = "path/to/your/inspection_log.xlsx"
```

## Troubleshooting

**Issue**: PDF doesn't load
- Make sure the PDF file is not corrupted
- Check if you have read permissions for the file

**Issue**: Excel file access error
- Close the Excel file if it's open in another program
- Check write permissions for the directory

**Issue**: Blurry PDF display
- Use the zoom controls (🔍+ / 🔍-) to adjust view
- The app renders at 2x resolution for clarity

## Tips

- Set the Cabinet ID before starting inspection to ensure proper logging
- Save annotations regularly to preserve your work
- Use the keyboard shortcuts for faster navigation (if implemented)
- The Excel file is automatically created on first run

## System Requirements

- Windows 10/11
- Python 3.8+
- 4GB RAM minimum
- Display resolution: 1280x720 or higher recommended

## License

Free to use for internal production and quality control purposes.
