# Excel to draw.io Converter - Specification

## 1. Project Overview

**Project Name:** Excel to draw.io Converter  
**Type:** Desktop GUI Application  
**Core Functionality:** Converts Excel shapes, diagrams, flowcharts, and cell-based diagrams into draw.io XML format  
**Target Users:** Business analysts, diagram creators, anyone who maintains diagrams in Excel but wants to export to draw.io

## 2. Features List

### 2.1 Core Features

| Feature | Description |
|---------|-------------|
| **File Selection** | Browse and select Excel files (.xlsx, .xls, .xlsm) via native file dialog |
| **Multi-Sheet Support** | Load and display all sheets from selected Excel file |
| **Sheet Selection** | Checkbox-based multi-select for sheets to convert |
| **Shape Conversion** | Convert Excel shapes (rectangles, ellipses, diamonds, connectors, arrows) to draw.io XML |
| **Cell-Based Diagram Support** | Convert merged cells with borders to shapes, preserving text content |
| **Styling Preservation** | Preserve fill colors, stroke colors, and text styling |
| **Connector Support** | Option to include/exclude connector lines and arrows |
| **Cell Color Support** | Option to include/exclude cell background colors |
| **draw.io Output** | Generate valid .drawio XML file |
| **Progress Indication** | Progress bar for long-running conversions |
| **Conversion Logging** | Real-time log of conversion activities |

### 2.2 GUI Features

| Feature | Description |
|---------|-------------|
| **Resizable Window** | Window can be resized by user (minimum 700x600, default 900x700) |
| **Scrollable Sheet List** | Sheet selection area with vertical scrollbar when many sheets exist |
| **Expandable Layout** | All expandable areas use `fill=tk.BOTH, expand=True` |
| **Paned Window Layout** | Use ttk.PanedWindow for resizable layout areas |
| **Responsive Components** | Components resize proportionally with window |
| **Log Area** | Scrollable text area showing conversion progress and messages |

## 3. UI/UX Specifications

### 3.1 Window Configuration

| Property | Value |
|----------|-------|
| **Default Size** | 900x700 pixels |
| **Minimum Size** | 700x600 pixels |
| **Resizable** | Yes (both width and height) |
| **Title** | "Excel to draw.io Converter" |
| **Window Style** | Native OS window decorations |

### 3.2 Layout Structure

```
+----------------------------------------------------------+
|  Title: "Excel to draw.io Converter"                     |
+----------------------------------------------------------+
|  [File Selection Frame]                                   |
|  | No file selected     [Browse Excel File] |           |
+----------------------------------------------------------+
|  [Paned Window - Vertical Split]                         |
|  +------------------------------------------------------+|
|  | Upper Paned Window (Horizontal Split)                 ||
|  | +------------------------+---------------------------+ ||
|  | | Sheet Selection        | Options Panel             |||
|  | | [Scrollable Frame]     | [ ] Include connectors    |||
|  | | [ ] Sheet1             | [ ] Include cell colors    |||
|  | | [ ] Sheet2             |                           |||
|  | | [ ] Sheet3             |                           |||
|  | | ...                    |                           |||
|  | +------------------------+---------------------------+ ||
|  +------------------------------------------------------+|
|  +------------------------------------------------------+|
|  | Lower Paned Window                                   ||
|  | +--------------------------------------------------+ ||
|  | | Log Area (ScrolledText)                          | ||
|  | |                                                  | ||
|  | +--------------------------------------------------+ ||
|  +------------------------------------------------------+|
+----------------------------------------------------------+
|  [Convert Button]        Status: X sheet(s) selected      |
+----------------------------------------------------------+
|  [Progress Bar]                                          |
+----------------------------------------------------------+
```

### 3.3 Component Specifications

#### File Selection Frame
- Height: ~50px fixed
- Contains: Label (file path), Browse button
- Background: Default frame background
- Padding: 10px vertical, 20px horizontal

#### Sheet Selection Panel
- Resizable width (shares horizontal space with options)
- Scrollable Canvas with vertical Scrollbar
- Default: Shows all sheets with checkboxes
- Checkbox behavior: Toggle selection, updates status

#### Options Panel
- Resizable width (shares horizontal space with sheets)
- Contains:
  - "Options:" label (bold)
  - Include connectors checkbox (default: checked)
  - Include cell colors checkbox (default: checked)
- Fixed minimum width: 200px

#### Log Area
- Expandable height (takes remaining vertical space)
- ScrolledText widget with vertical scrollbar
- Font: Courier 8pt
- Height: minimum 6 lines, expands with window
- State: disabled (read-only during operation)

#### Convert Button
- Size: 200x50px approximately
- Text: "Convert to draw.io"
- Color: Green (#28a745) background, white text
- State: DISABLED when no sheets selected, NORMAL otherwise

#### Progress Bar
- Mode: indeterminate
- Full width minus margins

#### Status Label
- Shows: "{n} sheet(s) selected" or "No sheets selected"
- Color: Blue text

### 3.4 Color Palette

| Element | Color |
|---------|-------|
| Primary Button | #667eea (blue-purple) |
| Convert Button | #28a745 (green) |
| Button Text | #FFFFFF (white) |
| Title | System default |
| Status Text | #0000FF (blue) |
| Log Area Background | System default |
| Log Area Font | Courier 8pt |

### 3.5 Spacing and Padding

| Element | Padding |
|---------|---------|
| Window | System default |
| Title Label | 10px vertical |
| File Frame | 10px vertical, 20px horizontal |
| Sheets Frame | 5px vertical/horizontal, 20px horizontal |
| Options Frame | 10px vertical, 20px horizontal |
| Convert Button | 10px vertical |
| Log Area | 10px all sides |
| Progress Bar | 5px vertical, 20px horizontal |
| Separators | 10px vertical, 20px horizontal |

## 4. User Workflow

### 4.1 Step-by-Step Conversion Flow

1. **Launch Application**
   - User runs `python gui_tkinter.py`
   - Window appears with default size 900x700
   - All controls visible but disabled (except Browse button)

2. **Select Excel File**
   - User clicks "Browse Excel File" button
   - Native file dialog opens
   - User selects .xlsx, .xls, or .xlsm file
   - File path appears in label
   - Application loads sheet names (shows progress bar during load)

3. **Review and Select Sheets**
   - All sheets displayed as checkboxes in scrollable list
   - All sheets selected by default
   - User can toggle individual sheets on/off
   - Status label updates with count

4. **Configure Options** (Optional)
   - User can uncheck "Include connectors/lines" if unwanted
   - User can uncheck "Include cell background colors" if unwanted

5. **Start Conversion**
   - User clicks "Convert to draw.io" button
   - Save dialog appears with default filename
   - User selects location and filename
   - Progress bar starts
   - Log area shows conversion progress

6. **Conversion Complete**
   - Success message box appears
   - Log shows completion message with file size
   - Application ready for next conversion

### 4.2 Error Handling Flow

1. **File Load Error**
   - Error message shown in dialog
   - Error logged to log area
   - Application returns to file selection state

2. **Conversion Error**
   - Error message shown in dialog
   - Error logged to log area
   - Convert button re-enabled
   - User can retry or select different file

3. **Cancelled Save**
   - Log shows "Conversion cancelled"
   - Application ready for next action

## 5. File Format Support

| Format | Extension | Description | Support Level |
|--------|-----------|-------------|---------------|
| Excel Workbook | .xlsx | Default Excel format (2007+) | ✅ Full |
| Excel Workbook | .xls | Legacy Excel format (97-2003) | ✅ Full |
| Excel Macro-Enabled | .xlsm | Excel with VBA macros | ✅ Full (macros ignored) |

## 6. Error Handling Specifications

| Error Type | Handling | User Feedback |
|------------|----------|---------------|
| Invalid file format | Catch exception, show messagebox | "Failed to load Excel file: {error}" |
| Corrupted file | Catch exception, show messagebox | "Failed to load Excel file: {error}" |
| Empty sheet | Skip sheet, log warning | "Warning: Sheet '{name}' is empty" |
| Conversion error | Catch exception, show messagebox | "Conversion failed: {error}" |
| Save cancelled | Return to ready state | "Conversion cancelled" |
| No sheets selected | Disable convert button | Status: "No sheets selected" |

## 7. Success Criteria

### 7.1 Functional Criteria

| ID | Criterion | Test Method |
|----|-----------|-------------|
| F1 | Window opens with 900x700 default size | Visual verification |
| F2 | Window can be resized by dragging | Drag window edge |
| F3 | Window has minimum size of 700x600 | Attempt to resize smaller |
| F4 | Browse button opens native file dialog | Click and verify |
| F5 | Excel file loads and displays sheet names | Select file, verify sheets |
| F6 | Sheet list is scrollable when >10 sheets | Load file with many sheets |
| F7 | Checkboxes toggle sheet selection | Click checkboxes, verify status |
| F8 | Options checkboxes work correctly | Toggle options, verify state |
| F9 | Convert button enables only when sheets selected | Verify state changes |
| F10 | Save dialog appears with correct default name | Click convert, verify dialog |
| F11 | Progress bar shows during conversion | Trigger conversion, verify |
| F12 | Log area shows conversion messages | Trigger conversion, verify |
| F13 | Success message shows after completion | Trigger conversion, verify |
| F14 | Error messages show on failure | Use invalid file, verify |

### 7.2 UI/UX Criteria

| ID | Criterion | Test Method |
|----|-----------|-------------|
| U1 | Sheet list expands to fill available space | Resize window, verify |
| U2 | Log area expands with window | Resize window, verify |
| U3 | Paned window allows resizing sections | Drag paned window sash |
| U4 | All text is visible (no overflow) | Load file, verify labels |
| U5 | Buttons have correct colors | Visual verification |
| U6 | Status updates reflect current selection | Toggle sheets, verify |

### 7.3 Compatibility Criteria

| ID | Criterion | Test Method |
|----|-----------|-------------|
| C1 | Works on Windows with Python 3.8+ | Run on Windows |
| C2 | Works on macOS with Python 3.8+ | Run on macOS |
| C3 | Works on Linux with Python 3.8+ | Run on Linux |
| C4 | No additional dependencies beyond requirements | Verify requirements.txt |
