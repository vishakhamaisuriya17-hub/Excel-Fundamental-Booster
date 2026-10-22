# Excel Fundamental Booster — PR. 1

A practical Microsoft Excel project focused on developing spreadsheet skills through formulas, logical functions, lookup operations, date calculations, dynamic references, and data analysis.

## Project Overview

**Excel Fundamental Booster** is a hands-on Excel project designed to strengthen fundamental Excel knowledge through practical exercises using student records, sales transactions, and employee information.

The project demonstrates how Excel functions can be used to organize data, automate calculations, apply logical conditions, retrieve information, analyze dates, and perform mathematical operations.

The workbook contains three primary datasets:

* Students Grade
* Sales Data
* Employee Data

Each dataset is used to practice specific Excel functions and solve practical data analysis problems.

## Project Objectives

The main objectives of this project are to:

* Understand relative and absolute cell references.
* Apply proper data entry and formatting techniques.
* Classify student performance using IF and nested IF formulas.
* Evaluate multiple conditions using AND and OR.
* Perform conditional counting, summation, and averaging.
* Retrieve information using VLOOKUP, XLOOKUP, INDEX, MATCH, and XMATCH.
* Manipulate text using Excel text functions.
* Calculate date differences and extract date components.
* Create dynamic cell references using INDIRECT and OFFSET.
* Apply mathematical rounding functions.
* Filter records based on specified conditions.
* Improve practical Excel and data analysis skills.

## Technologies and Tools

| Tool or Technology      | Purpose                                    |
| ----------------------- | ------------------------------------------ |
| Microsoft Excel         | Spreadsheet management and data analysis   |
| Excel Formulas          | Automated calculations and data processing |
| Logical Functions       | Conditional decisions and classification   |
| Lookup Functions        | Searching and retrieving information       |
| Date and Time Functions | Date calculations and analysis             |
| Text Functions          | Text extraction and transformation         |
| Dynamic Array Functions | Filtering and returning multiple records   |

## Workbook Structure

The project contains the following worksheets.

### 1. Project Instructions

Contains the project objectives, topics, essential tasks, and requirements for completing the exercises.

### 2. Students Grade

Contains student information and academic marks.

**Dataset columns:**

| Column          | Description                        |
| --------------- | ---------------------------------- |
| Student ID      | Unique identifier for each student |
| Name            | Student name                       |
| Math            | Mathematics marks                  |
| Science         | Science marks                      |
| English         | English marks                      |
| Enrollment Date | Student enrollment date            |

**Analysis performed:**

* Calculate average marks.
* Assign grades using conditional formulas.
* Count students who meet specific score requirements.
* Identify students scoring above a threshold in multiple subjects.
* Retrieve student names using student IDs.
* Extract and transform text.
* Analyze enrollment dates.
* Filter high-performing students.

### 3. Sales Data

Contains sales transaction information for different products, regions, and salespeople.

**Dataset columns:**

| Column      | Description                            |
| ----------- | -------------------------------------- |
| Sales ID    | Unique identifier for each transaction |
| Product     | Product sold                           |
| Region      | Sales region                           |
| Salesperson | Person responsible for the sale        |
| Amount      | Sales transaction amount               |
| Date        | Transaction date                       |

**Analysis performed:**

* Calculate total sales for a particular region and product.
* Apply conditional logic to sales transactions.
* Retrieve sales records using lookup functions.
* Calculate averages and rounded amounts.
* Identify product positions in a sales list.
* Apply fixed discount rates using absolute references.
* Demonstrate dynamic cell and range references.
* Perform date-based calculations.

**Important:** The Amount column represents transaction amounts, not necessarily individual product prices. A separate product price list would be required to retrieve actual product prices.

### 4. Employee Data

Contains employee information for salary analysis and date calculations.

**Dataset columns:**

| Column       | Description                         |
| ------------ | ----------------------------------- |
| Employee ID  | Unique identifier for each employee |
| Name         | Employee name                       |
| Department   | Employee department                 |
| Salary       | Employee salary                     |
| Joining Date | Employee joining date               |

**Analysis performed:**

* Retrieve employee salaries using XLOOKUP.
* Find employee departments using INDEX and MATCH.
* Calculate years of service.
* Calculate differences between dates.
* Extract years and months from dates.
* Transform employee names using text functions.
* Apply dynamic lookup formulas.

## Excel Topics Covered

### 1. Relative and Absolute Cell References

Learn how cell references behave when formulas are copied to other cells.

**Relative reference:**

```excel
=E2*10%
```

When copied down, the reference changes automatically.

**Absolute reference:**

```excel
=E2*$N$2
```

The reference `$N$2` remains fixed when the formula is copied.

**Key concepts:**

* Relative references
* Absolute references
* Mixed references
* Formula copying
* Data formatting

### 2. IF and Nested IF Functions

Use logical conditions to classify students and evaluate business rules.

**Example: Student grade classification**

```excel
=IF(G2>=90,"A+",IF(G2>=80,"A",IF(G2>=70,"B",IF(G2>=60,"C","F"))))
```

This formula assigns grades according to the student's average marks.

| Average Marks | Grade |
| ------------- | ----- |
| 90 and above  | A+    |
| 80–89.99      | A     |
| 70–79.99      | B     |
| 60–69.99      | C     |
| Below 60      | F     |

**Skills practiced:**

* Logical testing
* Nested conditions
* Grade classification
* Conditional calculations

### 3. IF with AND and OR

Evaluate multiple conditions in a single formula.

**AND example:**

Identify students scoring above 80 in both Math and Science.

```excel
=IF(AND(C2>80,D2>80),"Yes","No")
```

The result is Yes only when both conditions are satisfied.

**OR example:**

Determine whether a product qualifies for a discount when its amount exceeds 20,000 or another condition is satisfied.

```excel
=IF(OR(E2>20000,E2>=10000),"Eligible","Not Eligible")
```

In this example, the first condition is redundant because every amount above 20,000 is also above or equal to 10,000. It is included only to demonstrate OR syntax.

**Skills practiced:**

* Multiple-condition testing
* Logical decision-making
* Eligibility checking
* Conditional classification

### 4. COUNTIFS, SUMIFS, and AVERAGEIFS

Perform calculations based on specified criteria.

**COUNTIFS — Count students scoring above 60 in Math**

```excel
=COUNTIFS(C2:C21,">60")
```

**SUMIFS — Calculate sales for a particular region and product**

```excel
=SUMIFS(E2:E21,C2:C21,H2,B2:B21,I2)
```

Here:

* `E2:E21` contains the amounts to sum.
* `C2:C21` contains the regions.
* `H2` contains the selected region.
* `B2:B21` contains the products.
* `I2` contains the selected product.

**AVERAGEIF — Calculate the average sales amount above 30,000**

```excel
=AVERAGEIF(E2:E21,">30000",E2:E21)
```

**Skills practiced:**

* Conditional counting
* Conditional summation
* Conditional averaging
* Criteria-based analysis

### 5. VLOOKUP Function

Retrieve information by searching for a value in the first column of a table.

**Example: Retrieve a student name using Student ID**

```excel
=VLOOKUP(R2,A1:B21,2,FALSE)
```

The formula searches for the ID in `P5` and returns the corresponding student name.

**Skills practiced:**

* Exact-match lookup
* Student record retrieval
* Table-based searching
* Lookup range selection

### 6. INDEX and MATCH Functions

Combine INDEX and MATCH to retrieve information dynamically.

**Example: Retrieve an employee's department using Employee ID**

```excel
=INDEX(C2:C21,MATCH(L2,A2:A21,0))
```

MATCH identifies the position of the employee ID, and INDEX returns the corresponding department.

**Skills practiced:**

* Dynamic searching
* Position matching
* Flexible data retrieval
* Lookup without VLOOKUP's column-index restriction

### 7. Excel Text Functions

Manipulate and transform text values.

**LEFT and FIND — Extract the first word**

```excel
=LEFT(B2,FIND(" ",B2)-1)
```

**UPPER — Convert text to uppercase**

```excel
=UPPER(B2)
```

**LOWER — Convert text to lowercase**

```excel
=LOWER(B2)
```

**Skills practiced:**

* Text extraction
* Finding character positions
* Changing text case
* Basic text transformation

### 8. XLOOKUP Function

Retrieve matching information from separate lookup and return ranges.

**Example: Retrieve an employee's salary**

```excel
=XLOOKUP(J2,A2:A21,D2:D21,"employee not found")
```

The formula searches for the employee ID in column A and returns the matching salary from column D.

**Advantages:**

* Separate lookup and return ranges
* Flexible lookup direction
* Custom message when no match exists
* No requirement to specify a column index

### 9. XMATCH Function

Find the relative position of a value within a range.

**Example: Find the position of a product**

```excel
=XMATCH(P2,B2:B21,0)
```

The formula returns the position of the first exact match in the specified product range.

**Skills practiced:**

* Exact-match searching
* Position identification
* Dynamic lookup operations

### 10. INDIRECT Function

Convert a text value into a cell reference.

**Example: Retrieve a value from a referenced cell**

```excel
=INDIRECT(P2)
```

If cell P2 contains `B2`, the formula returns the value stored in B2.

**Skills practiced:**

* Dynamic cell referencing
* Text-based references
* Flexible formulas

### 11. OFFSET Function

Return a reference located a specified number of rows and columns from a starting cell.

**Example: Retrieve the next student's name**

```excel
=OFFSET(B2,1,0)
```

This formula moves one row down from B2 and returns the value from B3.

**Skills practiced:**

* Relative positioning
* Dynamic references
* Range-based calculations

### 12. Date and Time Functions

Perform date calculations and extract information from dates.

**YEAR — Extract the year**

```excel
=YEAR(F2)
```

**MONTH — Extract the month**

```excel
=MONTH(F2)
```

**TODAY — Return the current date**

```excel
=TODAY()
```

**DATEDIF — Calculate complete years between dates**

```excel
=DATEDIF(F2,TODAY(),"Y")
```

**Calculate the difference in days**

```excel
=H2-G2
```

**Skills practiced:**

* Date formatting
* Date component extraction
* Date differences
* Years of service calculations
* Date-based analysis

### 13. Mathematical Functions

Apply rounding functions to financial and numerical values.

**ROUND — Round to a specified number of decimal places**

```excel
=ROUND(E2,0)
```

**CEILING — Round up to the nearest multiple of 1,000**

```excel
=CEILING(E2,1000)
```

**FLOOR — Round down to the nearest multiple of 1,000**

```excel
=FLOOR(E2,1000)
```

**Skills practiced:**

* Numerical rounding
* Financial calculations
* Rounding to specified multiples
* Data presentation

### 14. FILTER Function

Extract records that satisfy a specified condition.

**Example: Filter students whose Math marks exceed 80**

```excel
=FILTER(students!A1:H21,students!G1:G21>80,"No students found")
```

This formula returns matching student records based on their avg marks.

**Alternative: Filter students whose math marks exceed 80**

```excel
=FILTER(Students!A1:H21,Students!c1:c21>80,"No students found")
```

**Skills practiced:**

* Criteria-based filtering
* Returning multiple records
* Dynamic array formulas
* Identifying high-performing students

## Practical Tasks Completed

The project focuses on the following practical exercises:

| No. | Task                                                | Excel Functions        |
| --- | --------------------------------------------------- | ---------------------- |
| 1   | Calculate student averages                          | AVERAGE                |
| 2   | Classify student grades                             | IF, Nested IF          |
| 3   | Count students above a score threshold              | COUNTIFS               |
| 4   | Identify students meeting multiple subject criteria | IF, AND                |
| 5   | Calculate regional and product-based sales          | SUMIFS                 |
| 6   | Calculate conditional averages                      | AVERAGEIF              |
| 7   | Retrieve student names using IDs                    | VLOOKUP                |
| 8   | Retrieve employee salaries                          | XLOOKUP                |
| 9   | Find employee departments                           | INDEX, MATCH           |
| 10  | Extract first words from text                       | LEFT, FIND             |
| 11  | Convert text to uppercase and lowercase             | UPPER, LOWER           |
| 12  | Find product positions                              | XMATCH                 |
| 13  | Create dynamic cell references                      | INDIRECT               |
| 14  | Create offset-based references                      | OFFSET                 |
| 15  | Calculate date differences                          | DATEDIF, TODAY         |
| 16  | Extract date components                             | YEAR, MONTH            |
| 17  | Apply rounding operations                           | ROUND, CEILING, FLOOR  |
| 18  | Filter records based on scores                      | FILTER                 |
| 19  | Apply relative and absolute references              | Cell references        |
| 20  | Format dates, numbers, and data                     | Excel formatting tools |

## How to Use This Project

Follow these steps to explore the workbook.

1. Clone or download the repository.
2. Open the Excel workbook in Microsoft Excel.
3. Navigate to the Project Instructions sheet.
4. Review the requirements for each task.
5. Open the Students Grade, Sales Data, and Employee Data sheets.
6. Examine the formulas and calculated results.
7. Change selected input values to understand how formulas respond.
8. Verify the calculated results against the original data.
9. Save your changes when finished.

### Clone the Repository

Run the following command in a terminal:

```bash
git clone https://github.com/vishakhamaisuriya17-hub/Excel-Fundamental-Booster.git
```

Navigate to the downloaded repository:

```bash
cd Excel-Fundamental-Booster
```

Open the workbook using Microsoft Excel.

## Software Compatibility

The workbook is designed for Microsoft Excel.

* Basic functions such as IF, COUNTIFS, SUMIFS, VLOOKUP, INDEX, MATCH, LEFT, FIND, UPPER, LOWER, YEAR, MONTH, ROUND, and FLOOR are supported by many Excel versions.
* XLOOKUP, XMATCH, and FILTER require a compatible modern version of Excel.
* DATEDIF, INDIRECT, and OFFSET can be used for date calculations and dynamic references in supported Excel versions.

Microsoft 365 or a recent compatible Excel version is recommended for the complete set of functions.

## Skills Demonstrated

### Technical Skills

* Microsoft Excel
* Formula creation and management
* Logical functions
* Conditional calculations
* Lookup and reference functions
* Text manipulation
* Date and time calculations
* Mathematical functions
* Dynamic array formulas
* Data formatting

### Analytical Skills

* Student performance analysis
* Sales data analysis
* Employee salary lookup
* Criteria-based data analysis
* Data organization
* Problem-solving
* Formula validation
* Logical reasoning

## Learning Outcomes

Through this project, I practiced using Excel to solve practical data-related problems.

Key learning outcomes include:

* Understanding how formulas automate repetitive calculations.
* Using conditional logic to classify and evaluate records.
* Applying multiple criteria to summarize data.
* Retrieving information from structured datasets.
* Manipulating text and extracting meaningful values.
* Performing date calculations and interpreting date differences.
* Using dynamic references to create flexible formulas.
* Filtering datasets to identify relevant records.
* Applying appropriate number and date formatting.
* Developing a stronger foundation for spreadsheet-based data analysis.

## Project Status

**Project:** PR. 1 Fundamental Booster

**Category:** Microsoft Excel and Data Analysis

**Level:** Beginner to Intermediate

**Status:** Practical Excel project

The workbook is intended for learning and demonstrating the practical application of Excel functions. Individual task completion should be verified against the formulas and results in the workbook.

## Author

**Vishakha Maisuriya**

GitHub: [vishakhamaisuriya17-hub](https://github.com/vishakhamaisuriya17-hub)

Repository: [Excel Fundamental Booster](https://github.com/vishakhamaisuriya17-hub/Excel-Fundamental-Booster)

---

This project demonstrates how Microsoft Excel can be used to organize information, automate calculations, retrieve records, and perform practical data analysis using built-in functions.
