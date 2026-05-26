# Microsoft Excel  Section 1: Understanding the Basics

A beginner-friendly guide to Excel fundamentals. Every concept is tied to the included dataset `employee_sales_data.xlsx` so you can follow along with real data.

---

##  Files in this repo

| File | Description |
|------|-------------|
| `employee_sales_data.xlsx` | Sample dataset — 10 employees with salary and sales data |

---

##  Dataset Overview

The dataset contains 10 employees across Sales, HR, and Finance with the following columns:

| Column | Description |
|--------|-------------|
| Emp ID | Unique employee identifier (E001–E010) |
| Name | Employee full name |
| Department | Sales, HR, or Finance |
| Join Date | Date of joining |
| Salary | Monthly salary |
| Jan / Feb / Mar Sales | Monthly sales figures |
| Total Sales | Auto-calculated using SUM |
| Avg Sales | Auto-calculated using IF + division |
| Commission (5%) | Auto-calculated at 5% of total sales |

---

##  Topics Covered

### 1. Understanding Excel Layout
The main components of the Excel workspace — Title Bar, Ribbon, Name Box, Formula Bar, Sheet Tabs, and Status Bar.

---

### 2. Workbook vs Worksheet
- **Workbook** = the entire `.xlsx` file
- **Worksheet** = a single tab inside the file

In the dataset, `employee_sales_data.xlsx` is the workbook and `Employee Sales Data` is one worksheet inside it.

---

### 3. Ribbons, Tabs, Groups, and Menus
The Ribbon is organised into **Tabs → Groups → Commands**.

Tabs used in this guide:
- `Home` — formatting (bold, fill colour, number format)
- `Formulas` — inserting SUM, AVERAGE, IF
- `Data` — sorting by department or salary
- `View` — freezing the header row

---

### 4. Customising the Quick Access Toolbar
Add frequently used commands above the Ribbon for one-click access.

> **Tip:** Add **Freeze Panes** to your QAT so you can lock Row 1 (the header) while scrolling through data.

---

### 5. Cells, Rows, and Columns

#### Naming and Selecting
Every cell has an address: **Column Letter + Row Number**

- `A1` → Header: Emp ID
- `B2` → Ananya Sharma (first employee)
- `E2:E11` → All 10 salary values
- `F2:H11` → Jan, Feb, Mar sales for all employees
- `I2:I11` → Total Sales (formula column)

#### Formatting
The header row (Row 1) in the dataset uses bold white text, a dark blue background fill, and number format `#,##0` for salary and sales columns.

#### Auto-Adjusting Column Width
If a cell shows `####`, the column is too narrow. Fix it by double-clicking the boundary between two column headers. Or select everything with `Ctrl + A` and double-click any boundary.

#### Creating and Deleting Rows / Columns
Right-click a row number → **Insert** to add a row above it. Right-click a column header → **Delete** to remove it.

---

### 6. Saving, Undo, and Redo

| Action | Shortcut |
|--------|----------|
| Save | `Ctrl + S` |
| Undo | `Ctrl + Z` |
| Redo | `Ctrl + Y` |
| Save As | `Ctrl + Shift + S` |

---

### 7. Intelligent Navigation

| Shortcut | Action |
|----------|--------|
| `Ctrl + End` | Jump to last used cell |
| `Ctrl + Home` | Jump back to A1 |
| `Ctrl + Arrow key` | Jump to last filled cell in that direction |
| `Ctrl + G` | Go To a specific cell |
| `Ctrl + F` | Find a value |

> In the dataset: `Ctrl + End` lands on the last cell of the summary section. `Ctrl + Down` from B2 jumps to B11 (last employee name).

---

### 8. Selecting Ranges

A range is written as `TopLeft:BottomRight` — for example `F2:H11`.

| Method | How |
|--------|-----|
| Select a range | Click and drag |
| Extend a selection | Hold `Shift` + click |
| Select non-adjacent cells | Hold `Ctrl` + click |
| Select to last used cell | `Ctrl + Shift + End` |
| Select entire column | Click the column header (A, B, C…) |

> To select all sales data in the dataset: click `F2`, hold `Shift`, click `H11`.

---

### 9. Copying and Pasting — Mastery

| Paste Option | Shortcut | Use case |
|---|---|---|
| Paste everything | `Ctrl + V` | General copy-paste |
| Paste Values only | `Alt + E + S + V` | Paste result without the formula |
| Paste Formats only | `Alt + E + S + T` | Copy styling to a new row |
| Paste Transpose | `Alt + E + S + E` | Flip rows and columns |

> Column I uses a SUM formula. If you paste it normally into another sheet, the formula breaks. Use **Paste Values** to paste just the number.

---

### 10. Basic Arithmetic Operations

#### Operators: +, -, *, /

All formulas start with `=`. Examples using the dataset:

| Formula | Result |
|---------|--------|
| `=F2+G2+H2` | Jan + Feb + Mar for Ananya = 54,000 |
| `=E2-E7` | Salary difference between Ananya and Neha = 10,000 |
| `=E6*0.1` | 10% of Vikram's salary = 4,000 |
| `=I6/3` | Vikram's monthly average ≈ 22,333 |

#### Built-in Functions

```excel
=SUM(F2:H2)
```
Total sales for one employee. Used in column I for all 10 rows.

```excel
=AVERAGE(E2:E11)
```
Average salary across all employees → 41,600

```excel
=MAX(E2:E11)
```
Highest salary → 55,000 (Karan Desai)

```excel
=MIN(E2:E11)
```
Lowest salary → 32,000 (Neha Joshi)

```excel
=COUNTA(A2:A11)
```
Count of employees → 10

```excel
=IF(I2=0, 0, I2/3)
```
Monthly average sales — returns 0 if the employee has no sales recorded.

---

### 11. Cell Referencing

| Type | Syntax | Behaviour when copied |
|---|---|---|
| Relative | `F2` | Both row and column adjust |
| Absolute | `$F$2` | Neither row nor column adjusts |
| Mixed (col locked) | `$F2` | Column F stays fixed, row adjusts |
| Mixed (row locked) | `F$2` | Row 2 stays fixed, column adjusts |

**Relative** — Used in column I of the dataset:
```excel
=SUM(F2:H2)
```
When copied down to row 3, it automatically becomes `=SUM(F3:H3)`.

**Absolute** — If the commission rate (5%) is stored in cell M1:
```excel
=I2*$M$1
```
Copying this formula down keeps `$M$1` fixed while `I2` adjusts to `I3`, `I4`, and so on.

**Mixed** — Useful in cross-tables or multiplication grids:
```excel
=$E2
```
Column E stays locked, row number adjusts freely.

---

### 12. Different Techniques of Fills

| Technique | How to use |
|---|---|
| Fill Handle | Drag the small green square at the bottom-right corner of a cell |
| Autofill Series | Type `1` and `2` in two cells, select both, drag the fill handle |
| Flash Fill | Start typing a pattern, press `Ctrl + E` |
| Fill Down | Select a cell + empty cells below, press `Ctrl + D` |
| Fill Right | Select a cell + empty cells to the right, press `Ctrl + R` |

```excel
Ctrl + D
```
Fill Down — copies the formula or value from the top cell into all selected cells below.

```excel
Ctrl + R
```
Fill Right — same as fill down but sideways.

```excel
Ctrl + E
```
Flash Fill — Excel detects your pattern and fills the rest automatically.

> Column I (Total Sales) was built by entering `=SUM(F2:H2)` in cell `I2`, then dragging the fill handle down to `I11`. Excel adjusted each row automatically.

---

## 📋 Dataset Quick Reference

| Emp ID | Name | Dept | Salary | Total Sales | Commission |
|--------|------|------|--------|-------------|------------|
| E001 | Ananya Sharma | Sales | 42,000 | 54,000 | 2,700 |
| E002 | Rohan Mehta | Sales | 38,000 | 42,000 | 2,100 |
| E003 | Priya Patel | HR | 45,000 | 0 | 0 |
| E004 | Arjun Nair | Sales | 35,000 | 33,500 | 1,675 |
| E005 | Sneha Gupta | Finance | 50,000 | 0 | 0 |
| E006 | Vikram Singh | Sales | 40,000 | 67,000 | 3,350 |
| E007 | Neha Joshi | Sales | 32,000 | 25,000 | 1,250 |
| E008 | Karan Desai | Finance | 55,000 | 0 | 0 |
| E009 | Meera Reddy | Sales | 36,000 | 39,000 | 1,950 |
| E010 | Suresh Kumar | HR | 43,000 | 0 | 0 |

---

##  Summary Formulas

These are already built into the dataset. Open the file and check the summary section at the bottom.

```excel
=COUNTA(A2:A11)
```
Total employees → 10

```excel
=SUM(I2:I11)
```
Grand total sales → 2,60,500

```excel
=MAX(E2:E11)
```
Highest salary → 55,000

```excel
=MIN(E2:E11)
```
Lowest salary → 32,000

```excel
=AVERAGE(E2:E11)
```
Average salary → 41,600

---

*Part of an ongoing Excel learning series. More sections coming soon.*
