# HR Data Cleaning Project (Excel)

Cleaning a messy HR dataset of 110,500 rows in Microsoft Excel: removing duplicates, handling missing values, standardizing text, converting data types, and reviewing salary outliers.

---

## 1. Project Overview

The raw dataset had duplicate employee records, inconsistent spellings, mixed date and salary formats, placeholder values, and many blanks. The goal was to turn it into a clean, analysis-ready dataset using Excel.

| Metric | Raw Data | Cleaned Data |
|---|---|---|
| Total rows | 110,500 | 100,001 |
| Unique EmployeeIDs | 100,000 (with 10,499 repeated rows) | 100,000, no repeats |
| Department values | 470 variants | 10 standard + "Check Manually" |
| City values | 114 variants | 8 standard + "Unknown" |
| Gender values | 11 variants | MALE / FEMALE / NOT SPECIFIED |
| PerformanceRating values | 9 variants (e.g. "five", "5.0", "N/A", blank) | 1 to 5 + "Not Rated" |
| Salary | Text with `$`, `Rs.`, commas, negatives | Numeric column `Salary_Clean` |
| JoiningDate | 8+ mixed formats | Proper Excel date (2010 to 2024) |

> The 100,001 rows in the cleaned file are 100,000 unique EmployeeIDs plus 1 row that had a blank EmployeeID in the raw data (kept as "Not Provided").

---

## 2. Dataset

**Raw file:** `HR_Messy_Dataset.xlsx` (sheet `RawData`, 110,500 rows x 11 columns)

**Raw columns:** EmployeeID, FullName, Department, City, Gender, JoiningDate, Salary, PhoneNumber, Email, PerformanceRating, Duplicate_Check

**Cleaned file columns:** EmployeeID, Full_Name, City, Joining Date, Day, Gender, Phone, Phone Validity Status, Department_rating (performance rating), Email, Salary_Clean, Department_Clear

### Data quality issues found in the raw data

- **Duplicates:** 10,499 repeated EmployeeIDs, 3,596 fully identical rows
- **Blanks:** about 1,000 blanks in every column, 500 rows with blank EmployeeID, 22,775 blank ratings
- **Department:** 470 raw variants (Sale/Sales, Procurment, Fiance, Marketting, Ops, CS, R&D/RnD, HR/H.R., I.T. etc.) plus hidden zero-width spaces
- **City:** 114 raw variants (Bangalore/Bengaluru, Bombay/Mumbai, Madras/Chennai, Calcutta/Kolkata, New Delhi/Delhi)
- **Gender:** 11 variants (M, F, male, FEMALE etc.)
- **Salary:** stored as text with `$`, `Rs.`, commas; 5,504 negative values; 5,372 "N/A" and 1,082 blank
- **JoiningDate:** many formats (2022-09-02, 10/30/2019, 25/05/22, 20 Apr 2013, May 01, 2011)
- **Phone:** different formats, some too short, and the placeholder `9999999999` (5,469 times)
- **Rating:** "five", "5.0", "N/A" alongside 1 to 5
- **Email:** `gmial.com` typos (21,811 rows)
- **Names:** extra spaces and inconsistent casing

---

## 3. Cleaning Process

### Step 1: Remove Duplicates
- Checked repeated EmployeeIDs (COUNTIF) and highlighted them (Conditional Formatting)
- Applied **Remove Duplicates** on EmployeeID
- Result: 110,500 rows reduced to 100,001; no repeated EmployeeID left

### Step 2: Handle Missing Values
- Counted blanks per column (COUNTBLANK) and located them (Go To Special, Blanks)
- Text fields filled with placeholders instead of guessing:
  - Name: "Not Provided"
  - City: "Unknown"
  - Gender: "NOT SPECIFIED"
  - Phone status: "Missing"
- Blank and "N/A" ratings converted to **"Not Rated"** (29,809 rows)

### Step 3: Text Cleaning
- Removed extra spaces and hidden characters (TRIM, CLEAN)
- Names and cities converted to Title Case (PROPER)
- Removed `$`, `Rs.` and `,` from salary (SUBSTITUTE)
- Checked phone number length (LEN) and created **Phone Validity Status**:
  - Valid: 94,032
  - Invalid - Incomplete: 4,979
  - Missing: 990
- Standardized city names (Bombay to Mumbai, Madras to Chennai, Calcutta to Kolkata, Bangalore to Bengaluru, New Delhi to Delhi)
- Standardized gender (M / male to MALE, F / female to FEMALE)
- Mapped 470 department variants to 10 standard departments using a mapping table and created the `Department_Clear` column. Rows that could not be mapped (blank department) are flagged **"Check Manually"** (1,576 rows)

### Step 4: Data Type Conversion
- Salary converted from text to number (`Salary_Clean`)
- JoiningDate converted from mixed formats to a proper date
- Added a `Day` column showing the weekday of the joining date
- Ratings written as "five" and "5.0" converted to the number 5

### Step 5: Outlier Review
- Checked salary range and found negative salaries (5,016 among unique employees)
- Negative salaries were set to blank

---

## 4. Verification

The cleaned file was compared against the raw file for every employee ID. All mappings were checked:

- Every raw city variant maps to the correct standard city
- Every raw gender variant maps to the correct standard value
- Every raw rating variant maps to the correct value ("five" and "5.0" to 5, "N/A" and blank to "Not Rated")
- All 100,000 raw EmployeeIDs are present in the cleaned data, none lost

---

## 5. Key Decisions

| Decision | Reason |
|---|---|
| Kept placeholder text ("Not Provided", "Unknown", "NOT SPECIFIED") for missing text fields | Guessing values for names, cities or gender would create false data |
| Used "Not Rated" for missing ratings | Ratings are a discrete 1 to 5 scale, so filling with an average would misrepresent employees |
| Used a mapping table for departments | 470 variants are too many to fix reliably by hand |
| Flagged unmapped departments as "Check Manually" | Keeps them visible for review instead of guessing |

---

## 6. Known Limitations

Items that are not yet cleaned in the final file:

| Issue | Count in cleaned data |
|---|---|
| Blank `Salary_Clean` (negative, "N/A" and blank salaries) | 10,168 |
| Salaries above 10 lakh (max 12,497,850; median 134,861), possible scale errors | 1,951 |
| Placeholder phone `9999999999` marked "Valid" | 4,937 |
| Emails with `gmial.com` typo | 19,859 |
| Emails with a space before `@` | 5,016 |
| Emails containing uppercase letters | 2,289 |
| Blank JoiningDate | 936 |
| Department flagged "Check Manually" | 1,576 |

Other notes:
- The department mapping table is still present in the workbook (columns P:Q)
- Negative salaries were blanked; they could also be sign typos (absolute value)

---

## 7. Author

**Your Name** | [LinkedIn](www.linkedin.com/in/aditya-s-verma-748645241) | [GitHub](https://github.com/adityasverma865-sys)
