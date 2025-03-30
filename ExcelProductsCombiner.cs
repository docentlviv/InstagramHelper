using System;
using System.Data;
using ExcelDataReader;
using ClosedXML.Excel;

namespace ExcelToTextConverter
{
    public class ExcelProductsCombiner{

        
        public string CombineWithProperties(string productsFilePath, string propertiesFilePath)
        {            
            string outputFilePath = $@"D:\InstagramExtract\Files\Result-{DateTime.Now.Ticks}.xlsx";

            try
            {
                // Register the encoding provider required by ExcelDataReader.
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                // Read the properties file and build a lookup dictionary keyed by 'Артикул'.
                DataTable propertiesTable = ReadExcelFile(propertiesFilePath);
                var propertiesLookup = propertiesTable.AsEnumerable().ToDictionary(
                        row => GetColumnValue(row, "Артикул"),
                        row => new
                        {
                            QnA = GetColumnValue(row, "Поширені запитання/Як використовувати(UA)"),
                            Dosage = GetColumnValue(row, "Застосування та дозування(UA)"),
                            Composition = GetColumnValue(row, "Аналітичний склад(UA)"),
                            Indication = GetColumnValue(row, "Призначення"),
                            ProductType = GetColumnValue(row, "Тип товару"),
                            AnimalSize = GetColumnValue(row, "Розмір тварини"),
                            Features = GetColumnValue(row, "Особливості")
                        }
                    );                    

                // Read the products file.
                DataTable productsTable = ReadExcelFile(productsFilePath);

                // Create a new workbook and add a worksheet using ClosedXML.
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Enriched Data");

                    // Write the header row.
                    int currentColumn = 1;
                    foreach (DataColumn dc in productsTable.Columns)
                    {
                        worksheet.Cell(1, currentColumn++).Value = dc.ColumnName;
                    }
                    // Append additional columns headers.
                    string[] extraColumns = {
                        "Поширені запитання/Як використовувати(UA)",
                        "Застосування та дозування(UA)",
                        "Аналітичний склад(UA)",
                        "Призначення",
                        "Тип товару",
                        "Розмір тварини",
                        "Особливості"
                    };

                    var columnsToAdd = extraColumns.Where(col => !productsTable.Columns.Contains(col)).ToList();
                    foreach (var col in columnsToAdd)
                    {
                        worksheet.Cell(1, currentColumn++).Value = col;
                    }

                    // Write the data rows.
                    int currentRow = 2;
                    foreach (DataRow productRow in productsTable.Rows)
                    {
                        currentColumn = 1;
                        // Write original product columns.
                        foreach (var item in productRow.ItemArray)
                        {
                            worksheet.Cell(currentRow, currentColumn++).Value = item.ToString();
                        }

                        // Use the common column 'Артикул' to find additional properties.
                        string artikul = GetColumnValue(productRow, "Артикул");
                        string qna = string.Empty, dosage = string.Empty, composition = string.Empty,
                               indication = string.Empty, productType = string.Empty, animalSize = string.Empty, features = string.Empty;
                        if (propertiesLookup.ContainsKey(artikul))
                        {
                            var prop = propertiesLookup[artikul];
                            qna = prop.QnA;
                            dosage = prop.Dosage;
                            composition = prop.Composition;
                            indication = prop.Indication;
                            productType = prop.ProductType;
                            animalSize = prop.AnimalSize;
                            features = prop.Features;
                        }
                        // Append extra properties.
                        worksheet.Cell(currentRow, currentColumn++).Value = qna;
                        worksheet.Cell(currentRow, currentColumn++).Value = dosage;
                        worksheet.Cell(currentRow, currentColumn++).Value = composition;
                        worksheet.Cell(currentRow, currentColumn++).Value = indication;
                        worksheet.Cell(currentRow, currentColumn++).Value = productType;
                        worksheet.Cell(currentRow, currentColumn++).Value = animalSize;
                        worksheet.Cell(currentRow, currentColumn++).Value = features;

                        currentRow++;
                    }

                    // Save the resulting workbook as an Excel file.
                    workbook.SaveAs(outputFilePath);
                }

                Console.WriteLine("Data enrichment complete. Result saved at: " + outputFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        
            return outputFilePath;    
        }

        // Helper method to read the first sheet of an Excel file into a DataTable.
        public System.Data.DataTable ReadExcelFile(string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Configure to use the first row as the header.
                    var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });
                    // Return the first worksheet.
                    return dataSet.Tables[0];
                }
            }
        }

        // Helper method to safely get a value from a DataRow based on a column name.
        static string GetColumnValue(DataRow row, string columnName)
        {
            return row.Table.Columns.Contains(columnName) ? row[columnName]?.ToString() ?? "" : "";
        }
    }
}