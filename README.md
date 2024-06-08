# Advanced Excel 

### Ribbon
- **MS OFFICE 2007+**
- Ribbon has Tabs: Home Tab, Insert Tab, etc.
- **Select Cursor:** For selection
- **Fill Cursor:** To fill the value, update the formula
- **Cell:** Value (Data) + Format of cell (Data Types)

### Excel File Structure
- **Excel File:** Workbook -> Book
- **Workbook:** Contains Worksheets -> Sheets
- **Keyboard Shortcuts:**
  - `Shift + F11` -> Insert new worksheet
  - `Ctrl + PgDn` -> Go to next sheet
  - `Ctrl + PgUp` -> Go to previous sheet
  - `Ctrl Key` -> Multiple selection
  - `Ctrl + 1` -> Open format cell option

### Basic Calculations in Excel
- `=C2+D2+E2+F2+G2`
- `=C2*D2`
- `=C2/D2`
- `=C2-D2`
- `=C2+D2-E2*F2/G2`

### Basic Functions in Excel
- **SUM, MAX, MIN, AVERAGE, COUNT:**
  - `=SUM(C2:G2)` -> Returns the addition of the given range
  - `=MAX(C2:G2)` -> Returns the maximum value from the given range
  - `=MIN(C2:G2)` -> Returns the minimum value from the given range
  - `=AVERAGE(C2:G2)` -> Returns the average of the given range
  - `=COUNT(C2:G2)` -> Counts how many cells in the given range contain numbers only

### Note
All above functions can accept single/multiple/combo ranges:
- `=SUM(C2:C7,E2:E6,C10,D10:F10)`
- `Alt + =` -> Apply the auto sum formula

## Day 2: 02-11-2023

### Advanced Calculations
- `Ctrl + Enter` -> Fill a common value over the selected range
- `Ctrl + !` -> Convert to number format with 2 decimal points
- `=LARGE(C2:G2,2)` -> Returns the 2nd largest value in the given range
- `=SMALL(C2:G2,2)` -> Returns the 2nd smallest value in the given range

### Counting with Criteria
- **COUNTIF:**
  - `=COUNTIF(C2:G2,">40")` -> Counts cells with values > 40
- **COUNTIFS:**
  - `=COUNTIFS(C2:G2,">=40",C2:G2,"<=80")` -> Counts cells with values between 40 and 80

### Cell References
1. **Relative Cell Reference:** Default cell address, e.g., `A2`, `B3:H55`
2. **Absolute Cell Reference:** Use `$` sign to lock the cell, e.g., `$A$2`, `$B$5:$H$66`
  - `F4` or `Fn + F4` -> Change cell reference types

### Text-Based Functions
- **Change Case:**
  - `=UPPER(B4)` -> Convert text to upper case
  - `=LOWER(B4)` -> Convert text to lower case
  - `=PROPER(B4)` -> Convert text to proper case
- **Concatenate Cells:**
  - `=B4&" "&C4`
  - `=B4&" "&C4&", "&D4`
- **Text to Columns:**
  - Split text from one column to multiple columns using Data Tab -> Data Tools -> Text to Column

## Day 3: 03-11-2023

### Flash Fill
- **MS Excel 2013+**
  - `Ctrl + E` -> AI-based flash fill

### Data Cleaning
- **Remove Extra Space:**
  - `=TRIM(A2)` -> Remove extra spaces
  - Copy the formula, go to source location, Paste Special (Ctrl + Alt + V) as value

### Logical Functions
- **IF, Nested IF:**
  - `=IF(D3>50,"Pass","Fail")`
  - `=IF(D3>90,"A",IF(D3>70,"B",IF(D3>50,"C","F")))`
- **AND, OR:**
  - `=AND(C4>=100,E4>=15000)`
  - `=IF(AND(C4>=100,E4>=15000),"yes","no")`
  - `=OR(C4>=100,E4>=15000)`
  - `=IF(OR(C4>=100,E4>=15000),"yes","no")`

### Data Validation
- Basics covered, more details to follow...

## Day 4: 06-11-2023

### Data Validation Continued
- **Addition with Condition:**
  - **SUMIF:**
    - `=SUMIF(E3:E16,"yes",D3:D16)` -> Total paid amount
    - `=SUMIF(E3:E16,"no",D3:D16)` -> Total unpaid amount
  - **SUMIFS:**
    - `=SUMIFS(D3:D16,B3:B16,"anurag",E3:E16,"yes")` -> Amount paid by Anurag

### Sorting and Filtering
- **Data Sorting:**
  - Ascending (A-Z)
  - Descending (Z-A)
  - Custom Order
  - Multilevel Data Sorting
- **Data Filter in Excel:**
  - `Ctrl + Shift + L`

## Day 5: 07-11-2023

### Filling Blank Cells
1. **Fill blank cells with a common message like "no data":**
  - Select the required range -> Home -> Find & Select -> Go to Special -> Blank -> OK -> Type 'no data' -> `Ctrl + Enter`
2. **Replace blank cells with nearby cell value:**
  - Select the required range -> Home -> Find & Select -> Go to Special -> Blank -> OK -> Type `=` then pass the reference of the nearby cell -> `Ctrl + Enter`

### Advanced Filter
- **Use of Advanced Filter:**
  - Fetch unique reports
  - Get unique lists of any item/field

### Data Consolidation
- **Data Validation:** Drop down list

## Day 6: 08-11-2023

### Date & Time Functions
- **Current Date and Time:**
  - `=TODAY()` -> Returns current system date
  - `Ctrl + ;` -> Returns current system date (fixed)
  - `=NOW()` -> Returns current system date with time
  - `Ctrl + :` -> Returns current system time (fixed)

### Date Functions
- **Extract Date Components:**
  - `=YEAR(F2)` -> Returns the year from the date
  - `=MONTH(F2)` -> Returns the month from the date
  - `=DAY(F2)` -> Returns the day from the date

### Calculate Past/Future Dates
- `=DATE(2023,11,8)` -> Returns the date based on given year, month, and day values
- `=DATE(YEAR(B4)+1,MONTH(B4)+9,DAY(B4)+20)` -> Returns the date after 1 year 9 months and 20 days

### Age Calculation
- `=DATEDIF(C4,$E$1,"y")` -> Returns the age in years
- `=DATEDIF(C4,TODAY(),"y")`
- `=DATEDIF(C4,TODAY(),"y")&" years"`
- `=DATEDIF(C4,$E$1,"y")&" years "&DATEDIF(C4,$E$1,"ym")&" months "&DATEDIF(C4,$E$1,"md")&" days"`

### Working Days Calculation
- `=NETWORKDAYS(A4,B4,F4:F7)` -> Total working days (Mon to Fri)
- `=NETWORKDAYS.INTL(A4,B4,11,F4:F7)` -> Total working days with custom weekends (Excel 2010+)

### Project End Date Calculation
- `=WORKDAY(A9,B9,F4:F7)` -> Project end date (Mon to Fri)
- `=WORKDAY.INTL(A9,B9,11,F4:F7)` -> Project end date with custom weekends (Excel 2010+)

## Day 7: 09-11-2023

### Pivot Tables
- **Keyboard Shortcut:** `Alt N V T`
- **Requirements:**
  - Category wise product total sales
  - Category wise product max, min, average, total sales
  - Category wise product total sales, % total sales
  - Category wise product total sales, % of parent total

### Pivot Table Design Tab
- **Report Layout:** Change the layout of Pivot Table
- **Subtotal**
- **Refresh Cache:** `Alt + F5`
- **Change Data Source**
- **Show Formula:** `Ctrl + `(~)`

## Day 8: 10-11-2023

### Pivot Table (continued)
- **Report Filter**
- **Slicer:** Visual filter (Excel 2010+)
- **Timeline:** Time-based filter (Excel 2013+)
- **Multiple Summarization:** 
  - Category wise product total sales with % of parent total and subtotal
- **Calculated Field:** Addition of new field
- **Calculated Item:** Addition of new item

### Pivot Table Customization
- **Design Tab:** 
  - Report Layout, Subtotals
  - % Running Total in, Difference From

### Lookup Functions
- **VLOOKUP:** `=VLOOKUP(lookup_value,table_array,col_index_num,range_lookup)`
- **HLOOKUP:** `=HLOOKUP(lookup_value,table_array,row_index_num,range_lookup)`

### Hyperlink
- **Insert Hyperlink:**
  - Insertion of a link to a different sheet, file, or web page

## Day 9: 13-11-2023

### More Functions
- **MATCH:** `=MATCH(lookup_value,lookup_array,match_type)`
- **INDEX:** `=INDEX(array,row_num,column_num)`
- **INDIRECT:** `=INDIRECT(ref_text,[a1])`

### Named Ranges
- **Create Named Range:** Name a cell range for easier reference in formulas

### Data Analysis ToolPak
- **Enabling Data Analysis ToolPak:**
  - File -> Options -> Add-ins -> Manage Excel Add-ins -> Go -> Check Data Analysis ToolPak -> OK

### Descriptive Statistics
- **Summary Statistics:**
  - Using Data Analysis ToolPak to generate summary statistics

## Day 10: 14-11-2023

### Advanced Charting
- **Creating Advanced Charts:** Combination charts, secondary axis, customizations
- **Using Sparklines:** Tiny charts within cells to show trends

### Power Query
- **Introduction to Power Query:** Data transformation and manipulation

### Macros
- **Recording Macros:** Automate repetitive tasks
- **VBA Basics:** Introduction to Visual Basic for Applications (VBA)

## Day 11: 15-11-2023

### Advanced Macros
- **Editing Recorded Macros:** Fine-tuning macro actions
- **Writing Simple VBA Scripts:** Customizing Excel behavior

### What-If Analysis
- **Scenario Manager:** Creating and managing different scenarios
- **Goal Seek:** Finding input values to achieve a desired result
- **Data Tables:** Creating one-variable and two-variable data tables

## Day 12: 16-11-2023

### Solver
- **Introduction to Solver:** Optimization tool for complex problems
- **Setting Up Solver:** Defining objective, constraints, and variables
- **Solver Examples:** Real-world problem-solving with Solver

### Advanced Power Query
- **Data Import and Transformation:** Advanced data import techniques
- **M Language Basics:** Introduction to Power Query's M language for advanced data manipulation

## Day 13: 17-11-2023

### Power Pivot
- **Introduction to Power Pivot:** Data modeling and large data sets
- **Creating Data Models:** Building relationships between tables
- **DAX Functions:** Introduction to Data Analysis Expressions (DAX) for advanced calculations

## Day 14: 20-11-2023

### Advanced Power Pivot
- **Complex DAX Functions:** Advanced calculations and measures
- **KPI Creation:** Defining and visualizing Key Performance Indicators (KPIs)
- **Using Power Pivot with Power BI:** Integrating Power Pivot models with Power BI

## Day 15: 21-11-2023

### Integration with Other Tools
- **Excel and PowerPoint:** Linking data and creating dynamic presentations
- **Excel and Word:** Mail merge and automated reporting
- **Excel and Access:** Importing and exporting data between Excel and Access

### Final Review and Q&A
- **Review of Topics Covered:** Recap of all key points and advanced techniques
- **Q&A Session:** Open floor for questions and additional help
