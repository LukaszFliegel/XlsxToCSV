using ExcelDataReader;
using System.Data;
using System.Text;

namespace XlsxToCSV
{
    internal class Program
    {
        //static ExcelConvert excelConvert = new ExcelConvert();

        static void Main(string[] args)
        {
            if (args.Length == 0)
                throw new ArgumentException("Missing path to excel file");

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using (var stream = File.Open(args[0], FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (data) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    StringBuilder csvStringBuilder = new();
                    foreach (var itemTable in result.Tables)
                    {
                        var table = itemTable as DataTable;

                        if (table == null)
                            throw new Exception("Error reading file");

                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            for (int j = 0; j < table.Columns.Count; j++)
                            {
                                var cell = "\"" + table.Rows[i].ItemArray[j] + "\";";
                                //Console.Write(cell);
                                csvStringBuilder.Append(cell);
                            }
                            //Console.WriteLine();
                            csvStringBuilder.AppendLine();
                        }
                    }

                    var csvFilePath = "temp.csv";
                    File.WriteAllText(csvFilePath, csvStringBuilder.ToString());

                    

                    //CsvToMt940Converter.ProcessCsvFile(csvFilePath);
                }
            }
        }
    }
}