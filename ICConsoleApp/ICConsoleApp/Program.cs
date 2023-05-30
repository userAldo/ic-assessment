using System.IO;
using System.Reflection;
using System;
using ICConsoleApp.Processes;

namespace ICConsoleApp
{
    class Program
    {
        private static string rootPath;

        static void Main(string[] args)
        {
            rootPath = Directory.GetParent(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)).Parent.FullName;
            XMLToRptConverter xmlToRptConverter = new XMLToRptConverter(rootPath);
            RptToMdbConverter rptToMdbConverter = new RptToMdbConverter(rootPath);
            MdbToCSVConverter mdbToCSVConverter = new MdbToCSVConverter(rootPath);

            while (true)
            {
                Console.WriteLine("Please select an option:");
                Console.WriteLine("1. XML To Rpt");
                Console.WriteLine("2. Rpt To MDB");
                Console.WriteLine("3. MDB to CSV");
                Console.WriteLine("0. Exit");
                Console.Write("Enter your choice: ");

                string input = Console.ReadLine();
                int choice;
                if (int.TryParse(input, out choice))
                {
                    switch (choice)
                    {
                        case 0:
                            Console.WriteLine("Exiting...");
                            return;
                        case 1:
                            Console.WriteLine("Converted XML to Rpt file");
                            try
                            {
                                xmlToRptConverter.ConvertXMLToRpt();
                                Console.WriteLine("XML converted successfully to Rpt");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Error converting XML to Rpt: " + ex.Message);
                            }
                            break;
                        case 2:
                            Console.WriteLine("Adding Rpt to mdb file");
                            try
                            {
                                rptToMdbConverter.ConvertRptToMdb();
                                Console.WriteLine("Rpt successfully added to mdb file");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Error adding Rpt to mdb file: " + ex.Message);
                            }
                            break;
                        case 3:
                            Console.WriteLine("Starting Mdb to CSV conversion");
                            try
                            {
                                mdbToCSVConverter.ConvertMdbToCSV();
                                Console.WriteLine("Mdb successfully converted to CSV");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Error converting Mdb to CSV: " + ex.Message);
                            }
                            break;
                        default:
                            Console.WriteLine("Invalid choice. Please try again.");
                            break;
                    }
                }
                else
                {
                    Console.WriteLine("Invalid input. Please enter a number.");
                }

                Console.WriteLine();
            }
        }
    } 
}
