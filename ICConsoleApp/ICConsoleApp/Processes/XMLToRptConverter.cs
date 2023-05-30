using System;
using System.Globalization;
using System.IO;
using System.Xml.Linq;
using System.Xml.XPath;

namespace ICConsoleApp.Processes
{
    class XMLToRptConverter
    {
        private string filePath;
        public XMLToRptConverter(String rootPath)
        {
           filePath = rootPath;
        }
        public void ConvertXMLToRpt()
        {
            // Read the contents of the BillFile.xml file
            string xmlFilePath = filePath + "\\Files\\BillFile.xml";
            XDocument xmldocument = XDocument.Load(xmlFilePath);

            // Parse XML
            string output = ParseXml(xmldocument);

            // output to export file
            string exportFilePath = filePath + "\\Files\\BillFile-" + DateTime.Now.ToString("MMddyyyy") + ".rpt";
            File.WriteAllText(exportFilePath, output);
            Console.WriteLine("Export File Created: " + exportFilePath);
        }

        private string GetElementText(XElement xmlElement, string xpath)
        {
            XElement element = xmlElement.XPathSelectElement(xpath);
            return element?.Value ?? string.Empty;
        }

        private string ParseXml(XDocument xDocument)
        {
            string output = "";

            int invoiceHeaderCount = 0;
            decimal billAmountSum = 0m;

            foreach (var billHeader in xDocument.Descendants("BILL_HEADER"))
            {
                string accountNumber = GetElementText(billHeader, "Account_No");
                string customerName = GetElementText(billHeader, "Customer_Name");
                string address1 = GetElementText(billHeader, "Address_Information/Mailing_Address_1");
                string address2 = GetElementText(billHeader, "Address_Information/Mailing_Address_2");
                string city = GetElementText(billHeader, "Address_Information/City");
                string state = GetElementText(billHeader, "Address_Information/State");
                string zip = GetElementText(billHeader, "Address_Information/Zip");
                string invoiceFormat = GetElementText(billHeader, "Bill/Bill_Tp");
                string invoiceNumber = GetElementText(billHeader, "Invoice_No");
                string billDate = DateTime.ParseExact(GetElementText(billHeader, "Bill_Dt"), "MMM-dd-yyyy", CultureInfo.InvariantCulture).ToString("MM/dd/yyyy");
                string dueDate = DateTime.ParseExact(GetElementText(billHeader, "Due_Dt"), "MMM-dd-yyyy", CultureInfo.InvariantCulture).ToString("MM/dd/yyyy");
                string firstNotificationDate = DateTime.ParseExact(billDate, "MM/dd/yyyy", CultureInfo.InvariantCulture)
                    .AddDays(5)
                    .ToString("MM/dd/yyyy");
                string secondNotificationDate = DateTime.ParseExact(dueDate, "MM/dd/yyyy", CultureInfo.InvariantCulture)
                    .AddDays(-3)
                    .ToString("MM/dd/yyyy");
                string balanceDue = GetElementText(billHeader, "Bill/Balance_Due");
                string currentDate = DateTime.Now.ToString("MM/dd/yyyy");
                string serviceAddress = GetElementText(billHeader, "Address_Information/Mailing_Address_1");

                output += $"AA~CT|BB~{accountNumber}|VV~{customerName}|CC~{address1}|DD~{address2}|EE~{city}|FF~{state}|GG~{zip}\n";
                output += $"HH~IH|II~R|JJ~{invoiceFormat}|KK~{invoiceNumber}|LL~{billDate}|MM~{dueDate}|NN~{balanceDue}|OO~{firstNotificationDate}|PP~{secondNotificationDate}|QQ~{balanceDue}|RR~{currentDate}|SS~{serviceAddress}\n";

                invoiceHeaderCount++;
                billAmountSum = decimal.Parse(balanceDue);
            }

            output = output.Insert(0, $"1~FR|2~8203ACC7-2094-43CC-8F7A-B8F19AA9BDA2|3~Sample UT file|4~{DateTime.Now.ToString("MM/dd/yyyy")}|5~{invoiceHeaderCount}|6~{billAmountSum.ToString(CultureInfo.InvariantCulture)}\n");

            return output;
        }
    }
}
