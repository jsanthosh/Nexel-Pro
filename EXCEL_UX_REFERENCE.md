# EXCEL UX BEHAVIOR REFERENCE

> **Purpose**: Exact UX behavior specifications for Microsoft Excel (365/2024 desktop),
> to be replicated in Nexel Pro. Every menu item, panel option, and interaction pattern
> is documented for implementation fidelity.

---

## 1. CHART ELEMENT SELECTION & FORMATTING

### 1.1 Clicking Chart Elements

When a chart is embedded in a worksheet, it starts as an **unselected object**. Interactions follow a two-level selection model:

**First click on chart**: Selects the chart as a whole object. A border with 8 resize handles appears. The ribbon switches to show the contextual "Chart Design" and "Format" tabs.

**Second click on a chart element**: Selects that specific element within the chart. Each element gets its own selection indicators:

| Element | Selection Indicator | Click Target |
|---------|-------------------|--------------|
| Chart Title | Blue bounding box with handles, cursor becomes text I-beam on double-click | The title text area |
| Axis Title | Blue bounding box with handles | The axis label text |
| Data Series | All data points in that series get selection handles (small squares on each bar/point/slice) | Any single bar, line segment, or pie slice |
| Individual Data Point | Single click on series selects series; second click on same point selects just that point | A specific bar/point/slice after series is selected |
| Legend | Bounding box with move/resize handles | The legend box |
| Legend Entry | Click legend selects legend; click specific entry selects that entry | Individual legend label |
| Axis (Value/Category) | Axis line and labels highlight | The axis line or any tick label |
| Gridlines (Major/Minor) | All gridlines of that type highlight | Any gridline |
| Plot Area | Dashed border around the plot region | The background area inside the axes |
| Chart Area | Selection on the outer chart boundary | The background outside the plot area |
| Data Labels | All labels for that series select; second click selects individual label | Any data label |
| Trendline | The trendline highlights with handles | The trendline itself |
| Error Bars | All error bars for that series highlight | Any error bar |

**Element dropdown**: In the "Format" contextual tab (and in the Format pane), there is a dropdown at the top-left labeled "Chart Elements" that lists every element in the current chart. Selecting from this dropdown selects that element. The dropdown always reflects the currently selected element.

### 1.2 Format Side Panel (Format Task Pane)

The Format pane opens on the right side of the window. It is triggered by:
- Double-clicking any chart element
- Right-click -> "Format [Element]..."
- Clicking "Format Selection" in the Format contextual tab
- Keyboard: Ctrl+1 when element is selected

**The pane title changes dynamically** based on the selected element:
- "Format Chart Title"
- "Format Data Series"
- "Format Axis"
- "Format Legend"
- "Format Plot Area"
- "Format Chart Area"
- "Format Data Labels"
- "Format Gridlines"
- "Format Trendline"
- "Format Error Bars"

**Pane structure**: The pane has icon tabs at the top, and each icon tab has sub-sections (accordion-style expandable groups):

#### Format Data Series (Bar/Column Chart)
Icon tabs:
1. **Fill & Line** (paint bucket icon)
   - Fill: No fill / Solid fill / Gradient fill / Picture or texture fill / Pattern fill / Automatic
     - Color picker, Transparency slider (0-100%)
     - For gradient: Type (Linear/Radial/Rectangular/Path), Direction preset, Gradient stops (add/remove/position/color/transparency/brightness)
   - Border: No line / Solid line / Gradient line / Automatic
     - Color, Width (pt), Dash type (Solid/Round Dot/Square Dot/Dash/Dash Dot/Long Dash/Long Dash Dot/Long Dash Dot Dot), Compound type, Cap type, Join type
   - Shadow (under Effects, but sometimes exposed here)

2. **Effects** (pentagon icon)
   - Shadow: Presets, Color, Transparency, Size, Blur, Angle, Distance
   - Glow: Presets, Color, Size, Transparency
   - Soft Edges: Size presets
   - 3-D Format: Top bevel, Bottom bevel, Depth, Contour, Surface material, Lighting
   - 3-D Rotation: Presets, X/Y/Z rotation, Perspective, Distance from ground

3. **Series Options** (bar chart icon) - CONTEXT SPECIFIC
   - Series Overlap: slider (-100% to 100%) — only for bar/column charts
   - Gap Width: slider (0% to 500%) — only for bar/column charts
   - Plot Series On: Primary Axis / Secondary Axis

#### Format Data Series (Pie Chart)
Icon tabs:
1. **Fill & Line** — same as above
2. **Effects** — same as above
3. **Series Options** (pie icon)
   - Angle of first slice: 0-360 degrees
   - Pie Explosion: 0-400% (how far slices pull apart)
   - For Doughnut: Doughnut Hole Size: 10-90%
   - NO Series Overlap or Gap Width (these are bar/column only)

#### Format Data Series (Line Chart)
Icon tabs:
1. **Fill & Line**
   - Line: same options as border above
   - Marker: marker options (Built-in types: circle/diamond/square/triangle/X/star/plus/dash, or None)
     - Marker Size, Fill, Border
   - NO "Fill" section (lines don't have fill, only line properties)
2. **Effects** — same
3. **Series Options**
   - Plot Series On: Primary/Secondary Axis
   - NO overlap/gap (those are bar/column only)

#### Format Axis
Icon tabs:
1. **Fill & Line** — for the axis line itself
2. **Effects** — shadows etc. on the axis
3. **Axis Options** (bar chart icon)
   - Bounds: Minimum (Auto or fixed), Maximum (Auto or fixed)
   - Units: Major (Auto or fixed), Minor (Auto or fixed)
   - Axis crosses: Automatic / Axis value / Maximum axis value
   - Display units: None / Hundreds / Thousands / Millions / Billions / Trillions
   - Logarithmic scale checkbox + Base field
   - Values in reverse order checkbox
   - Axis Type (for category axis): Automatically select / Text axis / Date axis
   - Vertical axis crosses: At category number / At maximum category / Automatically
   - Tick Marks: Major type (None/Inside/Outside/Cross), Minor type (same)
   - Labels: Label Position (Next to Axis / High / Low / None)
   - Number: Format Code, Category dropdown (same as cell number formatting)

#### Format Chart Title / Axis Title
Icon tabs:
1. **Title Options**
   - Fill & Line (for the text box background/border)
   - Effects
   - Size & Properties: Vertical alignment, Text direction, Autofit/resize, margins
2. **Text Options**
   - Text Fill & Outline: fill color for text, outline for text
   - Text Effects: shadow, reflection, glow on the text itself
   - Textbox: internal margins, autofit, columns

#### Format Legend
Icon tabs:
1. **Legend Options**
   - Fill & Line
   - Effects
   - Legend Options section: Legend Position (Top / Bottom / Left / Right / Top Right), Show legend without overlapping chart checkbox
2. **Text Options** — text fill, effects

#### Format Gridlines
Icon tabs:
1. **Fill & Line** — ONLY line options (color, width, dash type). No fill section.
2. **Effects** — shadow, glow (rarely used)
- Gridlines have NO dedicated "Gridline Options" section. They are simple lines.

#### Format Plot Area / Chart Area
Icon tabs:
1. **Fill & Line** — fill and border for the rectangular area
2. **Effects** — shadow, glow, soft edges, 3-D
3. **Size & Properties** (for Chart Area only) — fixed size or relative positioning

### 1.3 Chart Type-Specific Option Visibility

| Option | Bar/Column | Line | Pie/Doughnut | Scatter | Area | Radar |
|--------|-----------|------|-------------|---------|------|-------|
| Gap Width | Yes | No | No | No | No | No |
| Series Overlap | Yes | No | No | No | No | No |
| Angle of first slice | No | No | Yes | No | No | No |
| Pie Explosion | No | No | Yes (Pie only) | No | No | No |
| Doughnut Hole Size | No | No | Yes (Doughnut only) | No | No | No |
| Marker options | No | Yes | No | Yes | No | Yes |
| Line smoothing | No | Yes | No | Yes | No | No |
| Axis Bounds/Scale | Yes | Yes | No (no axes) | Yes | Yes | No |
| Drop Lines | No | Yes | No | No | Yes | No |
| High-Low Lines | No | Yes | No | No | No | No |
| Up/Down Bars | No | Yes (specific) | No | No | No | No |
| Trendline | Yes | Yes | No | Yes | Yes | No |
| Error Bars | Yes | Yes | No | Yes | Yes | No |

### 1.4 "Add Chart Element" Dropdown

Located in the **Chart Design** contextual tab, leftmost group. Contains:

1. **Axes** >
   - Primary Horizontal
   - Primary Vertical
   - (greyed/hidden for pie charts)
2. **Axis Titles** >
   - Primary Horizontal
   - Primary Vertical
3. **Chart Title** >
   - None
   - Above Chart
   - Centered Overlay
4. **Data Labels** >
   - None
   - Center
   - Inside End
   - Inside Base
   - Outside End
   - Data Callout
   - More Data Label Options...
5. **Data Table** >
   - None
   - With Legend Keys
   - No Legend Keys
   - More Data Table Options...
6. **Error Bars** >
   - None
   - Standard Error
   - Percentage (5%)
   - Standard Deviation
   - More Error Bar Options...
7. **Gridlines** >
   - Primary Major Horizontal
   - Primary Major Vertical
   - Primary Minor Horizontal
   - Primary Minor Vertical
   - More Gridline Options...
8. **Legend** >
   - None
   - Right
   - Top
   - Left
   - Bottom
   - More Legend Options...
9. **Lines** > (only for Line/Area charts)
   - None
   - Drop Lines
   - High-Low Lines
   - More Line Options...
10. **Trendline** >
    - None
    - Linear
    - Exponential
    - Linear Forecast
    - Moving Average
    - More Trendline Options...
11. **Up/Down Bars** > (only for Line charts with multiple series)
    - None
    - Up/Down Bars
    - More Up/Down Bar Options...

Each "More...Options" entry opens the Format pane for that element.

Checked items show a checkmark. Selecting "None" removes the element. Items are contextually hidden when they don't apply to the current chart type (e.g., Lines and Up/Down Bars hidden for bar charts, Axes hidden for pie charts).

### 1.5 Right-Click on Chart Elements

Right-clicking a chart element shows a context menu specific to that element. The menu always includes:

**Right-click on Data Series:**
1. (Mini toolbar: font, fill color) — floating above the menu
2. Cut
3. Copy
4. Paste Options (icons)
5. ---separator---
6. Reset to Match Style
7. Change Series Chart Type...
8. Select Data...
9. 3-D Rotation... (if applicable)
10. Add Data Labels
11. Add Trendline...
12. Format Data Series...

**Right-click on Chart Title:**
1. (Mini toolbar)
2. Cut / Copy / Paste
3. ---
4. Reset to Match Style
5. Font...
6. Edit Text
7. Delete Title
8. Format Chart Title...

**Right-click on Axis:**
1. (Mini toolbar)
2. Cut / Copy / Paste
3. ---
4. Reset to Match Style
5. Font...
6. Delete Axis
7. Format Axis...

**Right-click on Legend:**
1. (Mini toolbar)
2. Cut / Copy / Paste
3. ---
4. Reset to Match Style
5. Font...
6. Delete Legend
7. Format Legend...

**Right-click on Chart Area (background):**
1. (Mini toolbar)
2. Cut / Copy / Paste
3. ---
4. Reset to Match Style
5. Change Chart Type...
6. Select Data...
7. 3-D Rotation... (if applicable)
8. Move Chart...
9. Format Chart Area...

**Key pattern**: The last item is always "Format [Element Name]..." which opens the Format pane. "Delete [Element]" appears for removable elements. "Reset to Match Style" appears on all elements.

---

## 2. FORMAT CELLS DIALOG

Opened via: **Ctrl+1**, or right-click -> "Format Cells...", or Home tab -> Number group launcher (small diagonal arrow).

### 2.1 The 6 Tabs

1. **Number**
2. **Alignment**
3. **Font**
4. **Border**
5. **Fill**
6. **Protection**

### 2.2 Number Tab (Detailed)

**Layout**: Left side has a "Category" list, right side shows options specific to the selected category, top-right has a "Sample" preview showing how the active cell's value would appear.

**Categories and their options:**

1. **General**
   - No specific options
   - Description: "General format cells have no specific number format."
   - Sample shows the value as-is

2. **Number**
   - Decimal places: spinner (0-30, default 2)
   - Use 1000 Separator (,): checkbox
   - Negative numbers: listbox with 4 options:
     - -1234.10 (black)
     - 1234.10 (red, no minus)
     - (-1234.10) (black, parentheses)
     - (1234.10) (red, parentheses)

3. **Currency**
   - Decimal places: spinner (0-30, default 2)
   - Symbol: dropdown ($, EUR, GBP, JPY, none, etc. — long list, locale-dependent)
   - Negative numbers: same 4 styles as Number but with currency symbol

4. **Accounting**
   - Decimal places: spinner (0-30, default 2)
   - Symbol: dropdown (same as Currency)
   - Note: Accounting format aligns currency symbols at left edge and decimal points, shows zeros as dashes

5. **Date**
   - Type: listbox of date format patterns
     - 3/14/2012
     - 14-Mar
     - 14-Mar-12
     - Mar-12
     - March 14, 2012
     - March-12
     - M/D/YY
     - etc. (locale-dependent list, typically 15-20 options)
   - Locale (location): dropdown
   - Calendar type: dropdown (Western, Japanese, etc.)
   - Description explains the selected format

6. **Time**
   - Type: listbox of time format patterns
     - 1:30 PM
     - 1:30:55 PM
     - 13:30
     - 13:30:55
     - 3/14/12 1:30 PM
     - 37:30:55 (elapsed time)
     - etc.
   - Locale dropdown

7. **Percentage**
   - Decimal places: spinner (0-30, default 2)

8. **Fraction**
   - Type: listbox
     - Up to one digit (1/4)
     - Up to two digits (21/25)
     - Up to three digits (312/943)
     - As halves (1/2)
     - As quarters (2/4)
     - As eighths (4/8)
     - As sixteenths (8/16)
     - As tenths (3/10)
     - As hundredths (30/100)

9. **Scientific**
   - Decimal places: spinner (0-30, default 2)
   - Displays as e.g. "1.23E+04"

10. **Text**
    - No options
    - Description: "Text format cells are treated as text even when a number is in the cell. The cell is displayed exactly as entered."

11. **Special**
    - Type: listbox (locale-dependent)
      - Zip Code
      - Zip Code + 4
      - Phone Number
      - Social Security Number

12. **Custom**
    - Type: text field for entering custom format code
    - Listbox of predefined custom formats below (General, 0, 0.00, #,##0, #,##0.00, etc.)
    - Sample preview
    - "Delete" button to remove user-created custom formats
    - Description: "Type the number format code, using one of the existing codes as a starting point."

**Format code syntax** (shown in Custom category):
- `0` = required digit
- `#` = optional digit
- `,` = thousands separator
- `.` = decimal point
- `%` = multiply by 100 and show %
- `E+` / `E-` = scientific notation
- `;` separates positive;negative;zero;text sections
- `[Red]`, `[Blue]`, etc. for colors
- `"text"` for literal text
- `@` = text placeholder
- Date/time: `yyyy`, `mm`, `dd`, `hh`, `mm`, `ss`, `AM/PM`

### 2.3 Alignment Tab (Detailed)

**Text Alignment group:**
- Horizontal: dropdown
  - General (default — numbers right, text left)
  - Left (Indent)
  - Center
  - Right (Indent)
  - Fill
  - Justify
  - Center Across Selection
  - Distributed (Indent)
- Indent: spinner (0-250, only active when Horizontal is Left/Right/Distributed)
- Vertical: dropdown
  - Top
  - Center
  - Bottom (default)
  - Justify
  - Distributed
- Justify Distributed: checkbox (only active with Distributed alignment)

**Text Control group:**
- Wrap Text: checkbox — wraps text within the cell, expanding row height
- Shrink to Fit: checkbox — reduces font size so all content fits in cell width (mutually exclusive with Wrap Text)
- Merge Cells: checkbox — merges selected cells into one

**Right-to-Left group:**
- Text direction: dropdown
  - Context
  - Left-to-Right
  - Right-to-Left

**Orientation group:**
- Visual angled text box: a semicircular widget where you click to set angle
- Degrees: spinner (-90 to 90)
- Vertical text option: a separate tall-narrow box with "Text" written vertically (sets text to stack vertically, one letter per line)

### 2.4 Font Tab

- Font (family): listbox with preview (same list as ribbon font dropdown)
- Font Style: listbox (Regular, Italic, Bold, Bold Italic)
- Size: listbox/spinner (standard sizes 8-72, but any value can be typed)
- Underline: dropdown (None, Single, Double, Single Accounting, Double Accounting)
- Color: color picker dropdown (Theme Colors grid + Standard Colors + More Colors...)
- Normal Font: checkbox (resets to default workbook font)
- Effects group:
  - Strikethrough: checkbox
  - Superscript: checkbox
  - Subscript: checkbox
- Preview: shows sample text with current selections applied

### 2.5 Border Tab (Detailed)

**Layout**: Three main areas arranged from left to right and top to bottom.

**Presets group** (top):
Three large buttons:
1. **None** — removes all borders
2. **Outline** — applies border around the outside of selection only
3. **Inside** — applies borders only between cells (not outside edges)

A note says: "The selected border style can be applied by clicking the presets, preview diagram or the buttons above."

**Line group** (left side):
- **Style**: a listbox showing line styles:
  - None (blank)
  - Hair (thinnest, dotted)
  - Thin solid
  - Thin dashed (short dash)
  - Thin dash-dot
  - Thin dash-dot-dot
  - Medium solid
  - Medium dashed
  - Medium dash-dot
  - Medium dash-dot-dot
  - Thick solid
  - Double line
  - Slant dash-dot
  (13 styles total, shown as visual line samples)
- **Color**: dropdown color picker (same as font color picker)

**Border placement group** (right side):
- A **preview diagram** showing a rectangular cell (or merged area) with text "Text" inside
- **8 toggle buttons** around the preview for:
  - Top edge
  - Middle horizontal (inside horizontal border for multi-row selection)
  - Bottom edge
  - Left edge
  - Middle vertical (inside vertical border for multi-column selection)
  - Right edge
  - Diagonal down (\)
  - Diagonal up (/)
- Clicking a button or clicking directly on the corresponding edge in the preview toggles that border on/off with the currently selected line style and color
- Borders already applied show in the preview; you can click them again to remove

**Workflow**: User first selects a Line Style, then a Color, THEN clicks which edges to apply it to. This is the critical order — the style/color selection is "loaded" like a paint tool, then applied to edges.

When multiple cells are selected, the preview shows the "Inside" borders as well. If only one cell is selected, the middle-horizontal and middle-vertical buttons are greyed out.

### 2.6 Fill Tab

- **Background Color**: grid of theme colors + standard colors + "No Color" button + "More Colors..." button
- **Pattern Color**: dropdown color picker
- **Pattern Style**: dropdown showing 18 hatch/pattern options (none, 5% gray, 10% gray, 12.5% gray, 20% gray, 25% gray, 30% gray, 40% gray, 50% gray, 60% gray, 70% gray, 75% gray, 80% gray, 90% gray, Light diagonal stripe, Dark diagonal stripe, etc.)
- **Fill Effects...** button: opens sub-dialog with:
  - Gradient tab: One color / Two colors / Preset, Shading styles (Horizontal/Vertical/Diagonal Up/Diagonal Down/From Center/From Corner), Variants (4 direction options)
  - Pattern tab (same patterns as above but larger preview)
- **Sample** preview at bottom

### 2.7 Protection Tab

- **Locked**: checkbox (default ON for all cells)
  - Description: "Locking cells or hiding formulas has no effect until you protect the worksheet (Review tab, Protect Sheet button)."
- **Hidden**: checkbox (default OFF)
  - Description: "Hides the formula in the formula bar when the cell is selected. The value is still displayed in the cell."

---

## 3. RIGHT-CLICK CONTEXT MENUS

### 3.1 Right-Click on a Cell (or cell range)

A **mini toolbar** floats above the context menu (font, size, bold, italic, underline, fill color, font color, borders, merge, number format, percent, comma style, increase/decrease decimal, indent buttons).

**Context menu items** (in order):
1. Cut (Ctrl+X)
2. Copy (Ctrl+C)
3. Paste Options: (row of icons — Paste, Values, Formulas, Transpose, Formatting, Paste Link)
4. Paste Special... (Ctrl+Alt+V)
5. ---separator---
6. Smart Lookup (or Search)
7. ---separator---
8. Insert... (opens dialog: Shift cells right / Shift cells down / Entire row / Entire column)
9. Delete... (opens dialog: Shift cells left / Shift cells up / Entire row / Entire column)
10. Clear Contents (Delete key)
11. ---separator---
12. Quick Analysis (Ctrl+Q) — only when multiple cells with data are selected
13. Filter > (submenu with filter options if in a table/filtered range)
14. Sort > submenu:
    - Sort A to Z / Sort Z to A
    - Sort by Cell Color
    - Sort by Font Color
    - Sort by Cell Icon
    - Custom Sort...
15. ---separator---
16. Insert Comment (or New Note in newer versions)
17. New Note (or Edit Note if note exists)
18. ---separator---
19. Format Cells... (Ctrl+1)
20. Pick From Drop-down List... (Alt+Down Arrow)
21. ---separator---
22. Define Name...
23. Hyperlink... (Ctrl+K)

When a **single cell with data validation** is right-clicked, validation-related options may appear.

When cells are in a **Table**, additional items appear:
- Table > (submenu with Insert/Delete rows/columns, Select, Convert to Range)

### 3.2 Right-Click on a Row Header (row number)

Mini toolbar (same as cell).

**Context menu items:**
1. Cut
2. Copy
3. Paste Options (icons)
4. Paste Special...
5. ---separator---
6. Insert (inserts entire row(s) above — no dialog, immediate action)
7. Delete (deletes entire row(s) — no dialog, immediate action)
8. Clear Contents
9. ---separator---
10. Format Cells...
11. Row Height... (opens small dialog with height in points, e.g., default 15.00)
12. AutoFit Row Height
13. Hide (hides selected row(s))
14. Unhide (unhides rows; only shown if hidden rows are adjacent to selection)

### 3.3 Right-Click on a Column Header (column letter)

Mini toolbar (same as cell).

**Context menu items:**
1. Cut
2. Copy
3. Paste Options (icons)
4. Paste Special...
5. ---separator---
6. Insert (inserts entire column(s) to the left — immediate, no dialog)
7. Delete (deletes entire column(s) — immediate)
8. Clear Contents
9. ---separator---
10. Format Cells...
11. Column Width... (opens small dialog, default ~8.43 characters)
12. AutoFit Column Width
13. Hide
14. Unhide
15. ---separator---
16. (If in table context) Table-related options

### 3.4 Right-Click on a Sheet Tab

1. Insert... (opens template dialog for new sheet/chart sheet/macro sheet)
2. Delete (deletes the sheet, with confirmation if data exists)
3. Rename (makes tab name editable inline)
4. Move or Copy... (dialog to reorder or duplicate to another workbook)
5. View Code (opens VBA editor)
6. ---separator---
7. Protect Sheet...
8. Tab Color > (color picker submenu)
9. ---separator---
10. Hide (hides the sheet tab)
11. Unhide... (opens dialog listing hidden sheets)
12. ---separator---
13. Select All Sheets

### 3.5 Right-Click on a Chart (see Section 1.5 above)

---

## 4. TOOLBAR / RIBBON CONTEXT

### 4.1 Default Ribbon Tabs (no special selection)

1. **File** (backstage view, not a tab but a green/colored button)
2. **Home**
3. **Insert**
4. **Page Layout**
5. **Formulas**
6. **Data**
7. **Review**
8. **View**
9. **Help** (or Search)
10. **Developer** (optional, user-enabled)

### 4.2 Contextual Tabs When Chart Is Selected

When any chart is selected (clicked once), two additional tabs appear at the end of the ribbon, grouped under a header "Chart Tools" (in some versions):

1. **Chart Design** tab — contains:
   - Add Chart Element (dropdown — see Section 1.4)
   - Quick Layout (gallery of preset layouts)
   - Change Colors (gallery of color palettes)
   - Chart Styles gallery (visual style presets)
   - Switch Row/Column button
   - Select Data button (opens Select Data Source dialog)
   - Change Chart Type button (opens full chart type gallery)
   - Move Chart button (dialog: New sheet / Object in)

2. **Format** tab (contextual, for chart) — contains:
   - Current Selection group:
     - Chart Elements dropdown (lists all elements)
     - Format Selection button (opens Format pane)
     - Reset to Match Style button
   - Insert Shapes group: shape gallery, text box
   - Shape Styles group: shape fill, shape outline, shape effects galleries
   - WordArt Styles group: text fill, text outline, text effects
   - Arrange group: Bring Forward, Send Backward, Selection Pane, Align, Group, Rotate
   - Size group: Height/Width spinners

These tabs **disappear** when you click outside the chart (back into a cell).

### 4.3 Contextual Tabs When Picture/Image Is Selected

When an image is selected, one contextual tab appears, grouped under "Picture Tools":

1. **Picture Format** tab — contains:
   - Adjust group: Remove Background, Corrections (brightness/contrast), Color (saturation/tone/recolor), Artistic Effects, Compress Pictures, Change Picture, Reset Picture
   - Picture Styles group: gallery of preset frames/borders, Picture Border, Picture Effects (Shadow, Reflection, Glow, Soft Edges, Bevel, 3-D Rotation), Picture Layout (SmartArt conversion)
   - Accessibility group: Alt Text
   - Arrange group: Position, Wrap Text (not relevant in Excel), Bring Forward, Send Backward, Selection Pane, Align, Group, Rotate
   - Size group: Crop, Height/Width spinners
   - Crop button has dropdown: Crop, Crop to Shape, Aspect Ratio

### 4.4 Contextual Tab for Shapes

When a shape is selected, "Shape Format" tab appears:
- Insert Shapes group
- Shape Styles: fill, outline, effects galleries
- WordArt Styles: text fill, outline, effects
- Arrange group
- Size group

### 4.5 Contextual Tab for Tables (Structured Tables)

When cursor is inside a Table (Ctrl+T/Insert Table), "Table Design" tab appears:
- Properties group: Table Name, Resize Table
- Tools group: Summarize with PivotTable, Remove Duplicates, Convert to Range, Insert Slicer
- External Table Data group: Export, Refresh, Unlink (if external data)
- Table Style Options group: checkboxes for Header Row, Total Row, Banded Rows, First Column, Last Column, Banded Columns, Filter Button
- Table Styles gallery

### 4.6 Contextual Tab for Pivot Tables

When cursor is inside a PivotTable, two contextual tabs appear under "PivotTable Tools":
1. **PivotTable Analyze**: PivotTable Name, Options, Active Field, Group, Sort, Filter, Data source, Actions (Clear/Select/Move), Calculations, Tools (PivotChart, Timeline, Slicer), Show group
2. **Design**: Layout group (Subtotals, Grand Totals, Report Layout, Blank Rows), PivotTable Style Options (checkboxes), PivotTable Styles gallery

### 4.7 Home Tab State Reflection

The Home tab buttons **reflect the formatting state of the currently selected cell(s)**:

| Button | State Behavior |
|--------|---------------|
| **Bold (B)** | Appears pressed/highlighted when cell is bold |
| **Italic (I)** | Appears pressed/highlighted when cell is italic |
| **Underline (U)** | Appears pressed when cell has underline |
| **Strikethrough** | Appears pressed when cell has strikethrough |
| **Font Name** dropdown | Shows the font name of the current cell |
| **Font Size** dropdown | Shows the font size of the current cell |
| **Font Color** (A with colored bar) | The colored bar under the A matches the current font color |
| **Fill Color** (paint bucket) | The colored bar matches the current fill/background color |
| **Border** button | Icon shows the last-used border style |
| **Horizontal Alignment** (Left/Center/Right) | The matching alignment button appears pressed |
| **Vertical Alignment** (Top/Middle/Bottom) | The matching button appears pressed |
| **Wrap Text** | Appears pressed when wrap text is enabled |
| **Merge & Center** | Appears pressed when cell is part of a merged range |
| **Number Format** dropdown | Shows the category (General, Number, Currency, etc.) |
| **Percent Style (%)** | Appears pressed when format is percentage |
| **Comma Style (,)** | Appears pressed when comma/thousands format is active |
| **Increase/Decrease Decimal** | Always appear as normal buttons (no toggle state) |
| **Increase/Decrease Indent** | Always normal (no toggle state) |

When **multiple cells** with **mixed formatting** are selected:
- Toggle buttons (Bold, Italic, etc.) appear **unpressed** if mixed (some bold, some not)
- Font name shows the font of the **first cell** in the selection (top-left)
- Font size shows the size of the first cell, or is **blank** if sizes differ
- Font color bar shows the first cell's color

### 4.8 Formula Bar State

- Shows the **formula** (if cell has one) or the **raw value**
- For dates: shows the serial number or the formula, NOT the formatted date
- Name Box (left of formula bar) shows:
  - Cell reference (e.g., "A1") for single cell
  - Range (e.g., "A1:C5") when dragging to select
  - Named range name if a named range is selected
  - The name reverts to the active cell ref after Enter

---

## 5. STATUS BAR

### 5.1 Status Bar Layout (Left to Right)

The status bar runs along the bottom of the Excel window and has distinct zones:

**Left zone:**
- Mode indicator: displays "Ready", "Edit", "Enter", "Point", or "Calculate"
  - **Ready**: default state, no editing
  - **Edit**: cell is being edited (F2 or double-click)
  - **Enter**: user is typing into a cell (formula bar active)
  - **Point**: user is selecting cells for a formula reference
  - **Calculate**: workbook is recalculating (for large workbooks)

**Center zone:**
- Page number indicator (only in Page Break Preview or Print Preview mode)
- Macro recording indicator (red circle when recording)

**Right zone (aggregate calculations):**
- **Average**: shows average of selected numeric cells
- **Count**: shows count of non-empty selected cells
- **Sum**: shows sum of selected numeric cells
- These appear ONLY when 2+ cells are selected and at least one contains a number
- If only one cell is selected, this area is blank
- If only text cells are selected, only "Count" appears

**Far right zone:**
- View shortcut buttons: Normal, Page Layout, Page Break Preview (3 icons)
- Zoom slider: a slider with - / + buttons and percentage display
- Zoom percentage (e.g., "100%") — clicking it opens Zoom dialog

### 5.2 Real-Time Update Behavior

- The aggregate values (Sum, Average, Count) update **instantly** as the selection changes
- When you drag to extend a selection, values update in real-time during the drag
- When you Shift+Click to extend, values update immediately
- For large selections, there may be a brief delay with a "Calculate" indicator
- The values calculate from all selected cells, even across non-contiguous selections (Ctrl+Click)
- The zoom slider updates immediately during drag

### 5.3 Status Bar Right-Click Menu

Right-clicking the status bar shows a **toggle menu** — each item has a checkmark if visible. Items (in order):

**Cell Mode section:**
1. Cell Mode (Ready/Edit/Enter) — checked by default

**Signatures section:**
2. Signatures

**Information section:**
3. Permissions
4. Caps Lock (shows "Caps Lock" in status bar when active)
5. Num Lock
6. Scroll Lock (shows "SCROLL LOCK" indicator)
7. Fixed Decimal
8. Overtype Mode

**Selection statistics section:**
9. Average — checked by default
10. Count — checked by default
11. Numerical Count (counts only cells with numbers, vs Count which counts all non-empty)
12. Minimum (shows min of selected range)
13. Maximum (shows max of selected range)
14. Sum — checked by default

**View section:**
15. Upload Status
16. View Shortcuts — checked by default
17. Zoom Slider — checked by default
18. Zoom — checked by default (the percentage number)

**Default visible items**: Cell Mode, Average, Count, Sum, View Shortcuts, Zoom Slider, Zoom.

**Hidden by default but available**: Numerical Count, Minimum, Maximum, Caps Lock, Num Lock, Scroll Lock, Fixed Decimal, Overtype Mode, Signatures, Permissions.

### 5.4 Status Bar Calculation Details

| Statistic | What It Counts | Display Condition |
|-----------|---------------|-------------------|
| Average | Arithmetic mean of numeric cells | At least 1 numeric cell selected |
| Count | Number of non-empty cells (text + numbers + errors + booleans) | At least 1 non-empty cell |
| Numerical Count | Number of cells containing numbers only | At least 1 numeric cell |
| Sum | Sum of numeric cells | At least 1 numeric cell |
| Min | Smallest numeric value | At least 1 numeric cell |
| Max | Largest numeric value | At least 1 numeric cell |

- Cells with errors (#VALUE!, #REF!, etc.) are counted in Count but NOT in Average/Sum/Min/Max/Numerical Count
- Boolean TRUE = 1, FALSE = 0 for Sum/Average calculations
- Empty cells and text cells are ignored for numeric aggregations
- Date cells are treated as their numeric serial values

---

## 6. ADDITIONAL UX PATTERNS

### 6.1 Mini Toolbar

Excel shows a floating mini toolbar on:
- **Right-click** on cells (appears above the context menu)
- **Text selection** within a cell during edit mode

The mini toolbar contains (approximately):
Row 1: Font dropdown, Font Size, Increase Font Size, Decrease Font Size
Row 2: Bold, Italic, Underline, Borders dropdown, Fill Color, Font Color
Row 3: Merge & Center, Number Format dropdown, % style, Comma style, Increase Decimal, Decrease Decimal, Indent buttons

The mini toolbar **fades in** and **fades out** if the mouse moves away. Moving toward it makes it fully opaque.

### 6.2 Quick Access Toolbar (QAT)

Located above the ribbon (or optionally below). Default items:
- AutoSave toggle
- Save
- Undo
- Redo
- (Customizable — user can add any ribbon command)

### 6.3 Three Floating Buttons on Chart Selection

When a chart is selected, three small icon buttons appear floating to the upper-right of the chart:
1. **+** (Chart Elements) — checkbox list to toggle: Axes, Axis Titles, Chart Title, Data Labels, Data Table, Error Bars, Gridlines, Legend, Trendline. Each item has a right-arrow for sub-options.
2. **Paintbrush** (Chart Styles) — two sub-tabs: Style (gallery) and Color (color scheme gallery)
3. **Funnel** (Chart Filters) — two sub-tabs: Values (checkboxes to show/hide data series) and Names (checkboxes to show/hide categories)

### 6.4 Selection Highlighting Patterns

- **Active cell**: dark blue/green border, white fill (or cell's fill)
- **Selection range**: light blue/purple highlight over selected cells
- **Column/Row header highlight**: selected column letters / row numbers get a darker background color
- **Ctrl+Click**: creates non-contiguous selection, each range highlighted
- **Shift+Click**: extends selection from active cell to clicked cell
- **Entire row/column selection**: clicking header selects all cells; header itself goes bold/highlighted
- **Formula precedent/dependent arrows**: blue arrows pointing between cells

### 6.5 Auto-Complete / IntelliSense for Formulas

- When typing `=` followed by characters, a dropdown appears with matching function names
- Tab or click to accept the function name
- After `(`, a tooltip shows the function signature with parameter names
- Bold parameter indicates which parameter you're currently entering
- Ctrl+Shift+A inserts argument placeholders
