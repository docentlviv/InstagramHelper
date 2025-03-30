using System;
using System.IO;
using System.IO.Compression;
using HtmlAgilityPack; // Make sure to install the HtmlAgilityPack NuGet package

namespace HtmlTextExtractor
{
    public class InstagramMessagingExtract
    {
        public string ExtractMessages(string zipFilePath)
        {            
            // Create an extraction folder (named "Extracted") in the same directory as the ZIP file.
            string extractPath = Path.Combine(Path.GetDirectoryName(zipFilePath), "Extracted");
            // Define the output file path.
            string resultFilePath = $@"{zipFilePath}.txt";

            try
            {
                // Step 1: Unzip the file.
                if (Directory.Exists(extractPath))
                {
                    Directory.Delete(extractPath, true);
                }
                ZipFile.ExtractToDirectory(zipFilePath, extractPath);
                Console.WriteLine("Archive extracted to: " + extractPath);

                // If the result file exists, delete it to start fresh.
                if (File.Exists(resultFilePath))
                {
                    File.Delete(resultFilePath);
                }

                // Step 2: Recursively get all HTML files from the extracted folder.
                string[] htmlFiles = Directory.GetFiles(extractPath, "*.html", SearchOption.AllDirectories);

                // Step 3, 4 & 5: Read each HTML file, extract pure text, and append to the result file.
                using (StreamWriter writer = new StreamWriter(resultFilePath, append: false))
                {
                    foreach (string file in htmlFiles)
                    {
                        // Read the HTML file content.
                        string htmlContent = File.ReadAllText(file);

                        // Use HtmlAgilityPack to parse the HTML and extract only the text content.
                        HtmlDocument htmlDoc = new HtmlDocument();
                        htmlDoc.LoadHtml(htmlContent);
                        string plainText = htmlDoc.DocumentNode.InnerText.Trim();

                        // Write the extracted text and add a newline after processing each HTML file.
                        writer.WriteLine(plainText);
                    }
                }

                Console.WriteLine("Text extraction complete. The result is saved at: " + resultFilePath);

                return resultFilePath;
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }

            return resultFilePath;
        }
    }
}
