using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;


namespace ICConsoleApp.Processes
{
    class MdbToCSVConverter
    {
        private string filePath;
        public MdbToCSVConverter(String rootPath)
        {
            filePath = rootPath;
        }
        public void ConvertMdbToCSV()
        {
            string csvFilePath = filePath + "\\Files\\BillingReport.csv";
            string connectionString = $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={filePath}\Billing.mdb";

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                // Open the connection
                connection.Open();

                // Create a SQL query to retrieve the customers and bills data
                string csvQuery = "SELECT Customer.CustomerName, Customer.AccountNumber, Customer.CustomerAddress, Customer.CustomerCity, Customer.CustomerState, Customer.CustomerZip, " +
                               "Bills.BillDate, Bills.BillNumber, Bills.BillAmount, Bills.FormatGuid, Bills.AccountBalance, Bills.DueDate, Bills.ServiceAddress, Bills.FirstEmailDate, Bills.SecondEmailDate " +
                               "FROM Customer " +
                               "LEFT JOIN Bills ON Customer.ID = Bills.CustomerID";

                using (OleDbCommand command = new OleDbCommand(csvQuery, connection))
                using (OleDbDataReader reader = command.ExecuteReader())
                    if (reader.HasRows)
                    {
                        using (StreamWriter writer = new StreamWriter(csvFilePath))
                        {

                            // Write the header line to the CSV file
                            writer.WriteLine("CustomerName,AccountNumber,CustomerAddress,CustomerCity,CustomerState,CustomerZip,BillDate,BillNumber,BillAmount,FormatGuid,AccountBalance,DueDate,ServiceAddress,FirstEmailDate,SecondEmailDate");

                            // Iterate over the reader and write the data to the CSV file
                            while (reader.Read())
                            {
                                // Retrieve the values from the reader
                                string customerName = reader.GetString(0);
                                string accountNumber = reader.GetString(1);
                                string customerAddress = reader.GetString(2);
                                string customerCity = reader.GetString(3);
                                string customerState = reader.GetString(4);
                                string customerZip = reader.GetString(5);
                                string billDate = reader.GetDateTime(6).ToString("MM/dd/yyyy");
                                string billNumber = reader.GetString(7);
                                decimal billAmount = reader.GetDecimal(8);
                                string formatGuid = reader.GetString(9);
                                decimal accountBalance = reader.GetDecimal(10);
                                string dueDate = reader.GetDateTime(11).ToString("MM/dd/yyyy");
                                string serviceAddress = reader.GetString(12);
                                string firstEmailDate = reader.GetDateTime(13).ToString("MM/dd/yyyy");
                                string secondEmailDate = reader.GetDateTime(14).ToString("MM/dd/yyyy");

                                // Write the line to the CSV file
                                writer.WriteLine($"{customerName},{accountNumber},{customerAddress},{customerCity},{customerState},{customerZip},{billDate},{billNumber},{billAmount},{formatGuid},{accountBalance},{dueDate},{serviceAddress},{firstEmailDate},{secondEmailDate}");
                            }
                        }
                    }
                    else
                    {
                        throw new Exception("No data available for export.");
                    }
                }
                Console.WriteLine("Data exported to CSV file: " + csvFilePath);
            }
        }
    }

