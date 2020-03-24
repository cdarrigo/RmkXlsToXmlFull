using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using ExcelDataReader;
using Serilog;


// ReSharper disable CommentTypo
// ReSharper disable StringLiteralTypo
// ReSharper disable IdentifierTypo


namespace RmkXlsToXml
{
    /// <summary>
    /// Converts the incoming Remarketing XLS file to the output xml format. 
    /// </summary>
    public class RemarketingDataConverter
    {
        private readonly ILogger _logger;

        public RemarketingDataConverter(ILogger logger)
        {
            _logger = logger;
        }
        public bool ConvertRemarketingFile(ConverterConfiguration config)
        {
            //read the data from the source xls file to a list of strongly typed data.
            var data = ReadRmkDataFromSourceFile(config);
            // make sure we've got some data before continuing
            if (data == null)
            {
                _logger.Error("Error reading data from file.");
                return false;
            }

            if (!data.Any())
            {
                _logger.Error("No Remarketing data found in file.");
                return false;
            }

            // write the data to an Xml file
            WriteDataToXmlFile(config, data);
            return true;
        }

        private static string ComposeOutputFileName(ConverterConfiguration config)
        {
            // the output file will be the same name as the source file, but with an .xml file extension
            var sourceFileName = Path.GetFileNameWithoutExtension(config.SourceFile);
            var fullOutputFileName = Path.Combine(config.OutputPath, $"{sourceFileName}.xml");
            return fullOutputFileName;
        }
        /// <summary>
        /// Writes the Remarketing data to an xml file.
        /// </summary>
        private void WriteDataToXmlFile(ConverterConfiguration config, List<RemarketingData> data)
        {
            XmlWriterSettings settings = new XmlWriterSettings {Indent = true};
            var outputFileName = ComposeOutputFileName(config);
            if (File.Exists(outputFileName))
            {
                _logger.Warning($"Overwriting existing output file: {outputFileName}");
            }

            using (XmlWriter writer = XmlWriter.Create(outputFileName, settings))
            {
                writer.WriteStartElement("Remarketing"); // root node
                writer.WriteStartElement("FileInfo"); // FileInfo node
                writer.WriteElementString("RSAClientID", config.ClientConfiguration.RsaClientId);
                writer.WriteElementString("FileCreateDate", DateTime.Now.ToShortDateString());
                writer.WriteElementString("ItemCount", data.Count.ToString());
                writer.WriteEndElement(); // close FileInfoNode

                writer.WriteStartElement("RemarketingAssignmentList"); // starts the list of all the Remarketing assignments.
                foreach (var item in data)
                {
                    writer.WriteStartElement("RemarketingAssignment");
                    writer.WriteElementString("VIN",item.Vin);
                    writer.WriteElementString("AccountNumber", item.AccountNumber);
                    writer.WriteElementString("Year",item.Year);
                    writer.WriteElementString("Make",item.Make);
                    writer.WriteElementString("Model",item.Model);
                    writer.WriteElementString("Mileage",item.Mileage.ToString());
                    writer.WriteElementString("RepoDate",item.DateOfRepo.ToShortDateString());
                    writer.WriteElementString("ClearDate",item.DateOfClear.ToShortDateString());
                    writer.WriteElementString("LoanBalanceAmt",item.Balance.ToString(CultureInfo.InvariantCulture));
                    writer.WriteStartElement("VehicleLocationInfo");
                    writer.WriteElementString("IsVehicleAtCustomerSite","N");
                    writer.WriteElementString("LocationName",item.LocationOfUnit);
                    writer.WriteEndElement(); // closes VehicleLocationInfo
                    writer.WriteStartElement("CustomerInfo");
                    writer.WriteElementString("FullName", item.FullName);
                    writer.WriteEndElement(); // closes CustomerInfo
                    writer.WriteEndElement(); // closes RemarketingAssignment
                }
                writer.WriteEndElement(); // closes RemarketingAssignmentList
                writer.WriteEndElement(); // closes Remarketing
                writer.Flush();
            }
            _logger.Information($"Data has been written to: {outputFileName}");
        }
       
      

        public List<RemarketingData> ReadRmkDataFromSourceFile(ConverterConfiguration config)
        {
            var fileData = new List<RemarketingData>();

            // The incoming file is a BIFF 5.0 version of the excel file (Excel 2010)
            // so we have to use ExcelReader to read the contents.

            // running ExcelDataReading on .NET Core requires we specify the encoding provider
            
            // the easiest way to access the data is to read the entire workbook into a data set
            // and then parse the dataset data 
            DataSet ds;
            using (var stream = File.Open(config.SourceFile, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    ds = reader.AsDataSet();
                }
            }

            var colMap = config.ClientConfiguration.SourceColumnMap;
            // each worksheet in the excel file is a data table in the data set
            var sheet = ds.Tables[0];
            int rowIndex = -1;
            foreach(DataRow row in sheet.Rows)
            {
                // track the row number so we can skip the first N header rows
                rowIndex++;
                if (rowIndex < config.ClientConfiguration.NumberOfHeaderRows) continue;
                
                var items = row.ItemArray;
                // convert the data table data row to a strongly typed row of data.
                // ReSharper disable once UseObjectOrCollectionInitializer
                var data = new RemarketingData();

                
                data.AccountNumber = Get<string>(items, config.ClientConfiguration.SourceColumnMap.AccountNumber, rowIndex, nameof(data.AccountNumber));
                // if we encounter a missing account number, we'll use that to signify the end of the data rows.
                if (string.IsNullOrEmpty(data.AccountNumber))
                {
                    _logger.Debug($"encountered empty Account Number value at row {rowIndex} col {colMap.AccountNumber}. Assuming end of file.");
                    break;
                }

                data.LoanNumber = Get<string>(items, colMap.LoanNumber, rowIndex, nameof(data.LoanNumber));
                data.LastName = Get<string>(items, colMap.LastName, rowIndex, nameof(data.LastName));
                data.FirstName = Get<string>(items, colMap.FirstName, rowIndex, nameof(data.FirstName));
                data.Balance = Get<decimal>(items, colMap.LoanBalance, rowIndex, nameof(data.Balance)); 
                data.Year = Get<string>(items, colMap.Year, rowIndex, nameof(data.Year));
                data.Make = Get<string>(items, colMap.Make, rowIndex, nameof(data.Make));
                data.Model = Get<string>(items, colMap.Model, rowIndex, nameof(data.Model));
                data.Vin = Get<string>(items, colMap.Vin, rowIndex, nameof(data.Vin));
                data.Mileage = Get<int>(items, colMap.Mileage, rowIndex, nameof(data.Mileage));
                data.RepoAgentName = Get<string>(items, colMap.RepoAgentName, rowIndex, nameof(data.RepoAgentName));
                data.RepoAgentsLookup = Get<string>(items, colMap.RepoAgentsLookup, rowIndex, nameof(data.RepoAgentsLookup));
                data.LocationOfUnit = Get<string>(items, colMap.LocationOfUnit, rowIndex, nameof(data.LocationOfUnit));
                data.DateOfRepo = Get<DateTime>(items, colMap.DateOfRepo, rowIndex, nameof(data.DateOfRepo)); 
                data.DateOfClear = Get<DateTime>(items, colMap.DateOfClear, rowIndex, nameof(data.DateOfClear));
                fileData.Add(data);
            }
            _logger.Information($"Read {fileData.Count} row(s) from: {config.SourceFile}");
            return fileData;
        }


        private T Get<T>(object[] items, int? colIndex, int rowIndex, string propertyName, T defaultValue = default(T))
        {
            if (colIndex == null)
            {
                _logger.Information($"No source column mapping for property '{propertyName}'. Using default value:{ defaultValue?.ToString() ?? "null"}.");
                return defaultValue;
            }

            // Column indexes are articuled as 1 based, but they are read from a zero based collection
            var columnIndex = colIndex.Value - 1;
            var value = items[columnIndex];
            if (value == null) return defaultValue;
            try
            {
                return (T) Convert.ChangeType(value, typeof(T)) ;
            }
            catch (Exception e)
            {
                _logger.Error(e, $"Error attempting to convert source value to {typeof(T).Name}. Row # {rowIndex} Column # {colIndex} Value: {value}. Using default value:{ defaultValue?.ToString() ?? "null"}.");
                return defaultValue;
            }
        }

    }
}
