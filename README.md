# ic-assessment
# Running Solution
Using Visual Studio 2019
1. Select "Clone a Repository"
2. Under Repository location use the following link --- https://github.com/userAldo/ic-assessment.git  --- and clone
3. Double-Click ICConsoleApp.sln that is located in the Solution Explorer view
4. In the top menu select Build->Build Solution
5. Select Start Solution
6. Input 1 and press Enter to convert the BillFile.xml to an rpt
7. Input 2 and press Enter to add the rpt data to the Billing.mdb database
8. Input 3 and press Enter to export MDB to CSV 
9. Input 0 and Press Enter to exit the application
10. In solution explorer, press the refresh button ↻ - and the new files and reports should appear within the solution explorer under Files.

## XMLToRpt Key Functions
The code reads the contents of the XML file locaed at the path ('BillFile.xml') using XDocument.Load(). 
The code extracts elements from the XML document using XPath queries. the 'GetElementText() retrives the text values of specific elements in the XML. 
The code enforces that the Date Formatting be converted from the XML document to our desired format using 'DateTime.ParseExact()' and 'ToString'.
Then it creates a new file with the desired name "BillFile-(currentDate).rpt" and writes the generated output to it using 'File.WriteAllText()'.
The code begins to iterate over XML elements of type "BILL_HEADER" and extracts specific information from each element. 
and lastly it uses string manipulation by cocatenating various values together using string interpolation $ and attaching them to the 'output' string

### RptToMdbConverter
Reads ('BillFile-{currentDate}.rpt') and writes data to a Microsof Access database file ("Billing.mdb').
Establishes connection to Microsoft Access database using the OleDbConnection 
Retrieve Maximum IDs and executes SQL queries to retrieve the hieghest ID values from the "Bills" and "Customer" tables 
After it iterates over the lines read from the input file and processes each line it inserts data into the "Customer" and "Bills" table in the database and checks any existing records to avoid duplicates using SQL queries. 
Using parameterized Queries with placeholders to prevent SQL inhnjection and improve securitywhen inserting data into the database.
Then increments the highest cusomer ID and billID to ensure unique values when inserting new records.
Then the console ouput displays messages on the console for cases where dupliates ccustomer or bill records are encounterd. 

#### MdbToCSVConverter 
The code establishes a connnection to the Microsoft Access database using the 'OleDbConnection' class.
It creates an SQL quere to retrieve data from the 'Customer' and  'Bills' tables. and performs a left join operation to combine related data based on the 'CustomerID' column.
Retrieves data using OleDbDataReader when the code executes SQL query using the OleDbCommand.
Code creates CSV file (BillingReport.csv) using the StreamWriter and writes data to the file.
Code writers a header line containing the colum names to the CSV file.
Code iterates over the retrieved data to write each record a a line in the CSV file. It retrieves values fromm the OleDbDataReader based on the column index and formats them as needed. 
Lastly if no data is available for export, an exception is thrown with an appropriate error message. 

