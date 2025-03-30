using System;
using System.Data;
using System.IO;
using ExcelDataReader;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;

namespace ExcelToTextConverter
{
    public class ProductsFormatter
    {
        public string FormatProductsIntoTxt(string excelFilePath)
        {            
            string resultFilePath = $@"{excelFilePath}.txt";

            // Register the code pages provider (required for ExcelDataReader)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            try
            {
                // Read the first worksheet from the Excel file into a DataTable.
                DataTable table = ReadExcelFile(excelFilePath);

                // Define the list of columns to read.
                string[] columnsToRead = new string[]
                {
                    "Артикул",
                    "Название модификации (UA)",
                    "Бренд",
                    "Раздел",
                    "Цена",
                    "Ссылка",
                    "Описание товара (UA)",
                    "Короткое описание (UA)",
                    "Об`єм",
                    "Розмір",
                    "Вага",
                    "Поширені запитання/Як використовувати(UA)",
                    "Застосування та дозування(UA)",
                    "Розмір тварини",
                    "Призначення",
                    "Аналітичний склад(UA)",
                    "Особливості"
                };

                // Open the output text file for writing.
                using (StreamWriter writer = new StreamWriter(resultFilePath, append: false))
                {
                    // Process each row (product) in the Excel file.
                    foreach (DataRow row in table.Rows)
                    {
                        foreach (var colName in columnsToRead)
                        {
                            // Get the value if the column exists, otherwise return an empty string.
                            string value = GetColumnValue(row, colName);
                            writer.WriteLine($"{colName}: {value}");
                        }
                        // Add a blank line to separate records.
                        writer.WriteLine();
                    }
                }

                Console.WriteLine("Transformation complete. The result is saved at: " + resultFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }

            return resultFilePath;
        }

        // Reads the first worksheet of an Excel file into a DataTable.
        private DataTable ReadExcelFile(string filePath)
        {
            using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Use the first row as header.
                    var config = new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = tableReader => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    };

                    DataSet dataSet = reader.AsDataSet(config);
                    // Return the first worksheet.
                    return dataSet.Tables[0];
                }
            }
        }

        // Safely get a value from a DataRow for a given column name.
        private string GetColumnValue(DataRow row, string columnName)
        {
            return row.Table.Columns.Contains(columnName) ? row[columnName]?.ToString() ?? "" : "";
        }
    
        public void ConvertToPdf(string txtFilePath)
        {           
            string pdfFilePath = Path.ChangeExtension(txtFilePath, ".pdf");

            try
            {
                // Read the entire content of the text file.
                string content = File.ReadAllText(txtFilePath);

                // Create a new MigraDoc document.
                Document document = new Document();
                Section section = document.AddSection();

                // Optional: Set page margins.
                section.PageSetup.TopMargin = "2cm";
                section.PageSetup.BottomMargin = "2cm";
                section.PageSetup.LeftMargin = "2cm";
                section.PageSetup.RightMargin = "2cm";

                // Add a paragraph with the text file content.
                Paragraph paragraph = section.AddParagraph();
                paragraph.Format.Font.Name = "Verdana";
                paragraph.Format.Font.Size = 12;
                paragraph.AddText(content);

                // Render the document to PDF.
                PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true)
                {
                    Document = document
                };
                pdfRenderer.RenderDocument();
                pdfRenderer.PdfDocument.Save(pdfFilePath);                

                Console.WriteLine("Conversion complete. PDF saved at: " + pdfFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }

        }
    }
}
