# Student-Retention-Academic-Performance-Dashboard
An education analytics project using Excel for data cleaning and calculated fields, and Power BI for interactive analysis of student retention, academic performance, course-wise risk, and student support recommendations.

**Project Overview**

The project analyses 994 student records across 5 departments, 8 courses, and 3 years of study. The goal is to turn a raw student dataset into a cleaned analytical dataset and an interactive Power BI report that supports evidence-based review of academic outcomes.

**Tools Used**

- Microsoft Excel — data cleaning, validation, and calculated fields
- Power BI Desktop — Power Query transformations, DAX calculations, interactive slicers, KPI cards, charts, and report pages
  
**Methodology**

1. Reviewed the raw dataset for blanks, duplicate records, inconsistent category labels, data types, and invalid or inconsistent academic fields.
2. Cleaned and standardised the data in Excel, including checking student/course categories and recalculating Final Score (%), Grade, and Academic Status so the derived fields followed one consistent rule.
3. Imported the cleaned workbook into Power BI and verified data types and category values in Power Query. Text fields such as Course were trimmed/cleaned where needed to avoid duplicate-looking categories.
4. Created DAX calculated columns and measures for retention grouping, attendance bands, study-hour bands, early-warning flags, student counts, average scores, and promotion/non-promotion rates.
5. Built an interactive report with KPI cards, department/course/gender/state slicers, score comparisons, course-wise non-promotion rates, attendance-band analysis, and student distributions by department and year of study.
6. Added narrative, key-insight, and recommended-solution pages to explain the findings and outline possible academic support actions.
7. Validated displayed figures against the dataset and exported the report to PDF.

**Excel Formulas Used**

The cleaned workbook stores the assessment components in columns M:O, with the derived fields in columns P:R.
1. Final Score (%) — P2
Calculates the average of Assignment, Internal Exam, and Previous Exam scores only when all three values are numeric:
=IF(COUNT(M2:O2)<3,"",AVERAGE(M2:O2))
2. Grade — Q2
Assigns a grade based on the calculated final score:
=IF(P2="","",IF(P2>=90,"A+",IF(P2>=80,"A",IF(P2>=70,"B+",IF(P2>=60,"B",IF(P2>=50,"C",IF(P2>=40,"D","F")))))))
3. Academic Status — R2
Classifies students using the project’s score thresholds:
=IF(P2="","",IF(P2>=50,"Promoted",IF(P2>=40,"Needs Support","Repeat")))
These formulas were filled down the relevant rows. The thresholds above describe the rules used in this project; they should not be treated as universal academic grading standards.

**Power BI: DAX Calculations**

The model uses the main table Students Data. Examples of the core DAX calculations are shown below; field names must match those in the workbook/model.

**Calculated Columns**

Retention Status — groups Needs Support and Repeat as Not Retained for dashboard comparisons:
Retention Status =IF('Students Data'[Academic Status] = "Promoted","Retained","Not Retained")
Attendance Band — groups attendance values into ranges for comparison:
Attendance Band =SWITCH(TRUE(),'Students Data'[Attendance (%)] < 60, "Under 60%",'Students Data'[Attendance (%)] < 70, "60-69%",'Students Data'[Attendance (%)] < 85, "70-84%","85%+")
Study Hours Band — groups students by daily study time:
Study Hours Band =SWITCH(TRUE(),'Students Data'[Study Hours per Day] < 2, "Under 2h",'Students Data'[Study Hours per Day] < 3, "2-3h",'Students Data'[Study Hours per Day] < 4, "3-4h","4h+")
At Risk Flag — flags students below 50 in both Internal Exam and Assignment scores:
At Risk Flag =IF('Students Data'[Internal Exam Score (%)] < 50 &&'Students Data'[Assignment Score (%)] < 50,"At Risk","Not Flagged")

**Core Measures**

Total Students =DISTINCTCOUNT('Students Data'[Student ID])
Promoted Students =CALCULATE([Total Students],'Students Data'[Academic Status] = "Promoted")
Not Promoted Students =CALCULATE([Total Students],'Students Data'[Academic Status] IN {"Needs Support", "Repeat"})
Promotion Rate =DIVIDE([Promoted Students], [Total Students], 0)
Not Promoted Rate =DIVIDE([Not Promoted Students], [Total Students], 0)
Average Final Score =AVERAGE('Students Data'[Final Score (%)])
Additional measures calculate average Internal Exam, Assignment, Previous Exam, and Attendance values, and identify the number and proportion of at-risk students. Rates are formatted as percentages in Power BI.

**Dashboard Contents**

- Cards: Total Students, Not Promoted Students, Average Final Score, Promotion Rate
- Interactive slicers: Department, State, Course, and Gender
- Academic score comparison: Retained vs. Not Retained groups
- Course-wise Not Retained Rate
- Not Retained Rate by Attendance Band
- Student distribution by Department
- Student distribution by Year of Study
- Executive Narrative, Key Insights, and Recommended Solutions pages

**Key Findings**

In the current report, 890 of 994 students are promoted, while 104 are not promoted (93 Needs Support and 11 Repeat). The overall promotion rate is approximately 89.54%, and the average final score is approximately 64.83.

The report shows a substantial difference in average Internal Exam scores between promoted and non-promoted groups. Course-level rates are used to identify areas for further review, not as proof of course quality or teaching effectiveness. The early-warning rule is intended to prioritise support and does not identify every student who may need help.

**Notes and Limitations**
* Results depend on the cleaned dataset and the score/status rules documented above.
* Not Retained is a reporting group for students classified as Needs Support or Repeat; it does not mean these students have dropped out.
* The at-risk rule is a screening aid, not a definitive prediction of academic failure.

**Deliverables**
- Original Dataset
- Cleaned Excel Dataset
- Power BI File
- Power BI Report exported as PDF
- Project Documentation
