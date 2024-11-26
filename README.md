9. LEN Function
Syntax:

excel
Copy code
=LEN(text)
Example:
Find the length of the Name:

excel
Copy code
=LEN(A2)
10. SUBSTITUTE Function
Syntax:

excel
Copy code
=SUBSTITUTE(text, old_text, new_text, [instance_num])
Example:
Replace "HR" with "Human Resources" in the Department:

excel
Copy code
=SUBSTITUTE(C2, "HR", "Human Resources")
11. SEARCH Function
Syntax:

excel
Copy code
=SEARCH(find_text, within_text, [start_num])
Example:
Find the position of the text "Finance" in Department:

excel
Copy code
=SEARCH("Finance", C2)
12. ISNUMBER Function
Syntax:

excel
Copy code
=ISNUMBER(value)
Example:
Check if the value in Salary is a number:

excel
Copy code
=ISNUMBER(D2)
13. INDEX Function
Syntax:

excel
Copy code
=INDEX(array, row_num, [column_num])
Example:
Return the Salary of the employee in row 3:

excel
Copy code
=INDEX(D2:D6, 3)
14. MATCH Function
Syntax:

excel
Copy code
=MATCH(lookup_value, lookup_array, [match_type])
Example:
Find the position of Jane in the Name column:

excel
Copy code
=MATCH("Jane", A2:A6, 0)
15. UNIQUE Function (Excel 365 and Excel 2021)
Syntax:

excel
Copy code
=UNIQUE(array)
Example:
Return the unique departments in the dataset:

excel
Copy code
=UNIQUE(C2:C6)
16. IFS Function
Syntax:

excel
Copy code
=IFS(logical_test1, value_if_true1, logical_test2, value_if_true2, ...)
Example:
Assign job levels based on Age:

excel
Copy code
=IFS(B2<=30, "Junior", B2<=40, "Mid-Level", B2>40, "Senior")
17. COUNTIFS Function
Syntax:

excel
Copy code
=COUNTIFS(range1, criteria1, [range2], [criteria2], ...)
Example:
Count employees in the HR department with a Salary greater than 50,000:

excel
Copy code
=COUNTIFS(C2:C6, "", D2:D6, ">50000")
18. SUMIFS Function
Syntax:

excel
Copy code
=SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2], [criteria2], ...)
Example:
Sum the Salary for employees in the IT department with a Salary greater than 50,000:

excel
Copy code
=SUMIFS(D2:D6, C2:C6, "IT", D2:D6, ">50000")
19. AVERAGEIFS Function
Syntax:

excel
Copy code
=AVERAGEIFS(average_range, criteria_range1, criteria1, [criteria_range2], [criteria2], ...)
Example:
Calculate the average Salary for employees in the IT department:

excel
Copy code
=AVERAGEIFS(D2:D6, C2:C6, "IT")
20. TODAY Function
Syntax:

excel
Copy code
=TODAY()
Example:
Insert today's date:

excel
Copy code
=TODAY()
21. NOW Function
Syntax:

excel
Copy code
=NOW()
Example:
Insert the current date and time:

excel
Copy code
=NOW()
22. YEAR Function
Syntax:

excel
Copy code
=YEAR(date)
Example:
Extract the year from Join Date:

excel
Copy code
=YEAR(E2)
23. MONTH Function
Syntax:

excel
Copy code
=MONTH(date)
Example:
Extract the month from Join Date:

excel
Copy code
=MONTH(E2)
24. NETWORKDAYS Function
Syntax:

excel
Copy code
=NETWORKDAYS(start_date, end_date, [holidays])
Example:
Find the number of workdays between 01/01/2019 and 12/12/2020:

excel
Copy code
=NETWORKDAYS("01/01/2019", "12/12/2020")
25. EOMONTH Function
Syntax:

excel
Copy code
=EOMONTH(start_date, months)
Example:
Find the end of the month that is 2 months after 01/02/2021:

excel
Copy code
=EOMONTH("01/02/2021", 2)
26. FILTER Function (Excel 365 and Excel 2021)
Syntax:

excel
Copy code
=FILTER(array, include, [if_empty])
Example:
Filter employees with a Salary greater than 50,000:

excel
Copy code
=FILTER(A2:C6, D2:D6 > 50000)
27. FREQUENCY Function
Syntax:

excel
Copy code
=FREQUENCY(data_array, bins_array)
Example:
Find the frequency of salaries within specific ranges:

excel
Copy code
=FREQUENCY(D2:D6, {45000, 60000, 70000})
28. SEQUENCE Function (Excel 365 and Excel 2021)
Syntax:

excel
Copy code
=SEQUENCE(rows, [columns], [start], [step])
Example:
Create a sequence of numbers from 1 to 5:

excel
Copy code
=SEQUENCE(5,1,1,1)
29. RANDARRAY Function (Excel 365 and Excel 2021)
Syntax:

excel
Copy code
=RANDARRAY([rows], [columns], [min], [max], [integer])
Example:
Generate a 3x3 array of random numbers between 1 and 100:

excel
Copy code
=RANDARRAY(3, 3, 1, 100)
30. IFERROR Function
Syntax:

excel
Copy code
=IFERROR(value, value_if_error)
Example:
Return "Error" if division by zero occurs:

excel
Copy code
=IFERROR(D2/D3, "Error")
31. VLOOKUP Function
Syntax:

excel
Copy code
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
Example:
Find the Salary of Jane:

excel
Copy code
=VLOOKUP("Jane", A2:D6, 4, FALSE)
32. XLOOKUP Function (Excel 365 and Excel 2021)
Syntax:

excel
Copy code
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])
Example:
Find the Salary of Peter:

excel
Copy code
=XLOOKUP("Peter", A2:A6, D2:D6, "Not Found")
33. HLOOKUP Function
Syntax:

excel
Copy code
=HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])
Example:
Use HLOOKUP to find a salary in a horizontally structured dataset.

34. COUNT Function
Syntax:

excel
Copy code
=COUNT(value1, [value2], ...)
Example:
Count the number of numeric entries in the Age column:

excel
Copy code
=COUNT(B2:B6)
35. COUNTA Function
Syntax:

excel
Copy code
=COUNTA(value1, [value2], ...)
Example:
Count the number of non-empty cells in the Name column:

excel
Copy code
=COUNTA(A2:A6)
