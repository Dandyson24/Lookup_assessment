# Lookup and Pivot Table for basic analysis on Microsoft Excel

### Project overview
This project on Excel aims to implement lessons on lookup functions to extract, match and analyze data across multiple Excel sheets. The lookup functions are powerful formulae used to match data in two or more data sets. When combined with other functions and tools, they provide invaluable insight and ease of data analytics.
The second aim is to utilize pivot tables in summarizing data and providing quick insights for data driven decisions. For small datasets Microsoft Excel simplifies this process with her tools that provides fast and correct insights and are better than manual calculation/computation. The data provided was a small data set therefore Excel was appropriate for its cleaning and basic analytics using pivot tables.

### Objectives
1.	Understand the Dataset – The dataset consists of four different sheets containing structured data. Carefully review the information in each sheet. The first step in data analytics is to properly understand the dataset. Consequently, the Excel sheets with the tables were closely investigated. Both tables contain 20 rows and 10 columns. 
Data types include: Text, date and Numbers. 
Data is complete and consistent according to types and there were no duplicates therefore analyzable. 
2.	Apply Lookup Functions – Use VLOOKUP or XLOOKUP where applicable to retrieve relevant data based on the assignment questions and also the pivot assessment sheet.
3.	Answer the Assignment Questions – Follow the provided assignment sheet and complete the required tasks by applying lookup techniques effectively.

### Data sources
Data provided by 3MTT
### Tools
Microsoft Excel for matching and analytics using lookup functions and pivot tables.

### Data extraction and pivoting/analytics methodology
To identify duplicates: At the styles interface, conditional formatting was selected, followed by the highlight cell rules and then duplicate values. This was an essential step in the data cleaning process, no cells were highlighted therefore none was removed

To extract customer segment for a given Customer name in Table 1: The VLOOKUP function was deployed as directed and the exact customer segment was extracted following the conversion of the table to name ranges "last". =VLOOKUP($B$2,last,9,0). The VLOOKUP function can only look to the right and not towards the left. Therefore, was very appropriate for this.

To extract sales amount for a given date in Table 1:  Because the recommended VLOOKUP function does not look to the left of reference column, the data for sales was transposed to the right of the date and ascribed the name range "datesales".  The date column was formatted to DD/MM/YYY to conform with the question. The VLOOKUP function (=VLOOKUP($B$3,datesales,2,0)) was then deployed as directed and the exact sales amount for the given date was extracted.

To extract customer name for a given sales amount:  This was done using the XLOOKUP function. The  parameters were (lookup_value =  B4, lookup_array = name range “Sales “, return array= name range “ NAME“ and exact match to be returned. Plugging this into the formula gave:=XLOOKUP(B4,Sales,NAME,0). The customer name was extracted. XLOOKUP searches an array and returns the first exact match it finds. It is very appropriate here as it can lookup both data at the right and left of the lookup value.

To extract customer name for a given region:  This was done using the XLOOKUP function. The  parameters were (lookup_value =  B5, lookup_array = name range “Regions “, return array= name range “ NAME“ and exact match to be returned. Plugging this into the formula gave: =XLOOKUP(B5,Regions,NAME,0). The customer from the region was extracted. XLOOKUP searches an array and returns the first exact match it finds. It is very appropriate here as it can lookup both data at the right and left of the lookup value.

                                                         USE OF PIVOT TABLE FOR DATA ANALYSIS

PIVOT TABLE 1. Total order quantity for Liz Pelletier: Go to Home, insert, pivot table. On the Pivot table, the following were inserted into the fields; Row field: Customer name ; Sum of total sales in Values; Liz Pelletier was filtered and her total fund spent was collated. Also, the customer name can be placed in the filter field and when the filter arrow is clicked, tick select multiple items is ticked -> OK and Liz extracted.

PIVOT TABLE 2. Order priority category that produced the highest sales: On the Pivot table, Row field= Order priority; values = sum of total sales; Order priority was ordered from the largest to smallest at the editing, sort and filter pane(Z to A). The highest sales was extracted.

PIVOT TABLE 3. To extract 2009 sales: In the Pivot table Row field = Year; Values = Sum of sales; 2009 sales filtered. The other way is to place the year column header on the filter field of the pivot table and click the drop down on the year filter, tick select multiple items, OK-> select 2009, sales is extracted. In selecting the year, it is necessary to deselect the months and quarters unless these are necessary for the analysis.

PIVOT TABLE 4. Which region brought in the least profit for the year 2012? On the pivot table, Row field = Region; Values field = Sum of profit; Years in Filter field; Click on the markdown arrow on the filter, tick select multiple items, click OK and select 2012. The region with the least profit is extracted. 

PIVOT TABLE 5. Customer segment accounting for highest sales figure: Go to Home, insert, pivot table. On the Pivot table, the following were inserted into the fields: Row =Customer segment; Values = sum of sales, Sum of sales ordered from the largest to smallest on the editing/sort and filter pane, sort Z to A.

### Results and findings
1. Liz Lipellier bought 88 items

2. The high category of order priority produced the highest sales

3. 2009 sales stood at $11,741.707

4. The least profit in 2012 was from Ontario region

5. The home office customer segment accounted for the highest sales figure while the consumer segment was the least.

6. Muhammed MacIntyre belongs to the small business customer segment

7. The sales for 05/07/2009 was $2,484

8. The customer that spent $4,158.12 was Keith Dawkins

9. The customer from Yukon was Craig Yedwab

###  Recommendations
1. Helen Wasserman should be targeted and suggestions of goods and similar purchased goods made regularly.
2. High order priority goods should be prioritized while launching inquiries on why medium and low priority goods are not selling well might help. 
3.Reason behind a loss (negative profit) from Ontario region should be deciphered
4.Home office customer and corporate segments should be prioritized while support is given to the rest. Limitations
5. Further analysis on discount, shipping mode, profitability in different demographics, daily and monthly sales and revenue pattern, highest buyer and others will assist in improving sells for the store. These were outside the scope of the assignment

### Limitations
Inability of VLOOKUP to look up values that are at the left of the lookup value in a table
Excel 2019 does not support XLOOKUP. 
The use of Excel for data analytics may not be appropriate with larger volumes of data (big data). Other tools such as structured query language can be used in this case.
