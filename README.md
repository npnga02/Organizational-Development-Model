# AIESEC in Vietnam | National Organizational Development Model
## 📖Project Overview
### Objective
The Organizational Development Model (referred to as OD Model or ODM) is an index system (called OD Index or ODI), serving as a diagnosis and guidance definition tool for the national team and local executive boards to understand the current state and where to focus on to reach their ideal state. ODI is combined by 2 main indicators to define the overall status of an entity, which are:
Product Development Index (PDI) is focused on the portfolios we run
Health Development Index (HDI) is designed to analyze if we are managing our resources sustainably
This project aims to develop a comprehensive tool that helps LCs make data-driven decisions and strategically plan their development.

### Role
As the Data Manager, I was responsible for data collecting, data cleaning, performing calculations, visualization and building an interactive system readily available for decision-making. The project involved collaborating with various stakeholders, including the national board of directors and the local executive board.


## 🔮Approach and Process
### Data Collection
There are two methods for data collection:
With exchange functions: Data is automatically extracted from EXPA, with reference to the code from https://github.com/AIESEC-LK/extract-expa-applications-to-sheets, and additional necessary metrics are included.
With non-exchange functions: Data is collected manually from the commission workspaces.

### Data Preparation
Once the data is collected, the next step is to prepare it for analysis. Key tasks include:
Creating data validation in spreadsheets to ensure values match perfectly, facilitating future calculations.
Structuring the OD database in an unpivot table format to support the entire OD process.
Calculating exchange indicators from raw application data and inputting the calculated performance results into the OD database as shown below:

Adjusting local data at certain points of time due to inactive operations.
Performing calculations based on intermediate metrics collected.
	For example: %Sign-ups to Youth-engaged = #Youth-engaged / #Sign-ups, where #Youth-engaged and #Sign-ups are primarily collected.

### Data Input
To create the PDI and HDI tables, the following data processing steps are performed:

Use data validation for the metrics in row 7 to ensure consistency and facilitate future formula creation.
Create formulas to retrieve the weight value of parameters from the "OD Model | Metrics" tab to the PDI/HDI tabs of functions. For example, cell G8: 
=INDEX('Data Validation'!$C:$E,MATCH($A7 & G$7,'Data Validation'!$C:$C & 'Data Validation'!$D:$D,),3)
In row 9, if a parameter has a standard number, create a formula to fetch that value from the "OD Model | Metrics" tab. For example, cell R9
=INDEX('Data Validation'!$C:$F,MATCH($A7 & Q$7,'Data Validation'!$C:$C & 'Data Validation'!$D:$D,),4)
If the benchmark is the highest entity, retrieve the maximum value, such as in cell H9.
=MAX(G10:G18)

Create formulas to look up values in a database based on multiple criteria. For example, cell G10
=INDEX(Database,
MATCH($B$4 & $B10,Database!$A:$A & Database!$B:$B,),
MATCH($A$7 & G$7,Database!$1:$1 & Database!$2:$2,))

Normalize values to a common scale, such as in cell H10 
=IFERROR(IF(G10/H$9*10>10,10,
IF(G10/H$9*10<0,0,G10/H$9*10)))

Calculate individual values for the More, Better, and Sustainable groups. For example, cell F10.
=SUMPRODUCT((H$19:L$19="x")*IFERROR(H10:L10*1, 0), G$8:K$8)

Sum the values of More, Better, and Sustainable groups and then create a ranking in columns D and E.

### Data Processing
The component scores for PDI and HDI of various functions are consolidated in the "OD Model | Total" tab where the total PDI and total HDI are calculated. These totals are then averaged to derive the ODI, which is used to rank the Clusters.

###Data Output & Interpretation
The data is visualized using a bubble chart, where the PDI and HDI are represented on the two axes, and the ODI is depicted by the size of the bubbles. The clusters are divided into segments based on their scores: less than 2.5, less than 5, less than 7.5, and from 7.5 to 10.

