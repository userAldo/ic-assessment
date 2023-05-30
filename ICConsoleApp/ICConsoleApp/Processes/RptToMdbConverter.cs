using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ICConsoleApp.Processes
{
    class RptToMdbConverter

    {
        private string filePath;
        public RptToMdbConverter(String rootPath)
        {
           filePath = rootPath;
        }
        public void ConvertRptToMdb()
        {
            string rptFilePath = filePath + "\\Files\\BillFile-" + DateTime.Now.ToString("MMddyyyy") + ".rpt";
            string connectionString = $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={filePath}\Billing.mdb";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                // Open the connection
                connection.Open();

                int highestBillId = 0;
                int highestCustomerId = 0;

                // Get the highest ID from the Bills table
                string maxBillIdQuery = "SELECT MAX(ID) FROM Bills";
                using (OleDbCommand maxBillIdCommand = new OleDbCommand(maxBillIdQuery, connection))
                {
                    var result = maxBillIdCommand.ExecuteScalar();
                    highestBillId = result != null && result != DBNull.Value ? (int)result : 0;
                }

                // Get the highest CustomerID from the Customer table
                string maxCustomerIdQuery = "SELECT MAX(ID) FROM Customer";
                using (OleDbCommand maxCustomerIdCommand = new OleDbCommand(maxCustomerIdQuery, connection))
                {
                    var result = maxCustomerIdCommand.ExecuteScalar();
                    highestCustomerId = result != null && result != DBNull.Value ? Convert.ToInt32(result) : 0;
                }

                string[] lines = System.IO.File.ReadAllLines(rptFilePath);
                // Iterate over the lines and insert data into the corresponding tables and fields
                foreach (string line in lines)
                {
                    // Split the line by the pipe character to get individual fields
                    string[] fields = line.Split('|');

                    //Line starting with "1~" is for Header and will be skipped for adding to table
                    if (fields[0].Substring(0, 2) == "1~")
                    {
                        continue;
                    }

                    //Line starting in AA indicates Customer
                    if (fields[0].Substring(0, 2) == "AA")
                    {
                        int count = 0;

                        //Look for Account Number in Table to avoid duplicate customer addition
                        string accountNumberExistsQuery = "SELECT COUNT(*) FROM Customer WHERE AccountNumber = @AccountNumber";

                        using (OleDbCommand accountNumberExistsCommand = new OleDbCommand(accountNumberExistsQuery, connection))
                        {
                            accountNumberExistsCommand.Parameters.AddWithValue("@AccountNumber", fields[1].Substring(3));

                            // Execute the query to check if the account number exists
                            count = (int)accountNumberExistsCommand.ExecuteScalar();
                        }

                        //If Account Number exists in db skip customer addition
                        if (count > 0)
                        {
                            Console.WriteLine("Customer with account number {0} already exists in database", fields[1].Substring(3));
                            continue;
                        }

                        //Increase over previously highest customerID to insert new customer 
                        highestCustomerId++;
                        string customerQuery = "INSERT INTO Customer (ID,CustomerName, AccountNumber, CustomerAddress, CustomerCity, CustomerState, CustomerZip) " + "VALUES (@ID ,@CustomerName, @AccountNumber, @CustomerAddress, @CustomerCity, @CustomerState, @CustomerZip)";

                        //Add New Customer
                        using (OleDbCommand command = new OleDbCommand(customerQuery, connection))
                        {
                            command.Parameters.AddWithValue("@ID", highestCustomerId);
                            command.Parameters.AddWithValue("@CustomerName", fields[2].Substring(3).Replace(",", ""));
                            command.Parameters.AddWithValue("@AccountNumber", fields[1].Substring(3));
                            command.Parameters.AddWithValue("@CustomerAddress", fields[3].Substring(3));
                            command.Parameters.AddWithValue("@CustomerCity", fields[5].Substring(3));
                            command.Parameters.AddWithValue("@CustomerState", fields[6].Substring(3));
                            command.Parameters.AddWithValue("@CustomerZip", fields[7].Substring(3));
                            command.ExecuteNonQuery();
                        }
                    }

                    //Line starting in HH indicates a Bill
                    if (fields[0].Substring(0, 2) == "HH")
                    {
                        int count = 0;

                        //Look for BillNumber in Bills Table to avoid duplicate bill addition
                        string billNumberExistsQuery = "SELECT COUNT(*) FROM Bills WHERE BillNumber = @BillNumber";

                        using (OleDbCommand billNumberExistsCommand = new OleDbCommand(billNumberExistsQuery, connection))
                        {
                            billNumberExistsCommand.Parameters.AddWithValue("@BillNumber", fields[3].Substring(3));

                            // Execute the query to check if the bill number exists
                            count = (int)billNumberExistsCommand.ExecuteScalar();
                        }

                        //If BillNumber exists in db skip bill addition
                        if (count > 0)
                        {
                            Console.WriteLine("Bill with bill number {0} already exists in database", fields[3].Substring(3));
                            continue;
                        }

                        //Increase over highest BillID
                        highestBillId++;
                        string billQuery = "INSERT INTO Bills (ID, BillDate, BillNumber, BillAmount, FormatGuid, AccountBalance, DueDate, ServiceAddress, FirstEmailDate, SecondEmailDate, DateAdded, CustomerID) " + "VALUES (@ID, @BillDate, @BillNumber, @BillAmount, @FormatGuid, @AccountBalance, @DueDate, @ServiceAddress, @FirstEmailDate, @SecondEmailDate, @DateAdded, @CustomerID)";

                        using (OleDbCommand command = new OleDbCommand(billQuery, connection))
                        {
                            // Add parameter values
                            command.Parameters.AddWithValue("@ID", OleDbType.Integer).Value = highestBillId;
                            command.Parameters.AddWithValue("@BillDate", fields[4].Substring(3));
                            command.Parameters.AddWithValue("@BillNumber", fields[3].Substring(3));
                            command.Parameters.AddWithValue("@BillAmount", fields[6].Substring(3));
                            command.Parameters.AddWithValue("@FormatGuid", fields[1].Substring(3));
                            command.Parameters.AddWithValue("@AccountBalance", fields[9].Substring(3));
                            command.Parameters.AddWithValue("@DueDate", fields[5].Substring(3));
                            command.Parameters.AddWithValue("@ServiceAddress", fields[11].Substring(3));
                            command.Parameters.AddWithValue("@FirstEmailDate", fields[7].Substring(3));
                            command.Parameters.AddWithValue("@SecondEmailDate", fields[8].Substring(3));
                            command.Parameters.AddWithValue("@DateAdded", DateTime.Now.ToString("MM/dd/yyyy"));
                            command.Parameters.AddWithValue("@CustomerID", highestCustomerId);

                            // Execute the query
                            command.ExecuteNonQuery();
                        }

                    }

                }

            }
        }

    }
}
