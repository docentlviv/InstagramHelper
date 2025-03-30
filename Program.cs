using System;
using System.IO;
using System.IO.Compression;
using ExcelToTextConverter;
using HtmlAgilityPack; // Make sure to install the HtmlAgilityPack NuGet package

namespace HtmlTextExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            // first step add properties like FAQ and ingridients to the resulting file
            var productscombiner = new ExcelProductsCombiner( );

            string resultPath = @"D:\InstagramExtract\Files\website_export_raw.xlsx";
            resultPath = productscombiner.CombineWithProperties(resultPath, @"D:\InstagramExtract\Files\specification_cosmetics.xlsx");

            resultPath = productscombiner.CombineWithProperties(resultPath, @"D:\InstagramExtract\Files\specification_toy.xlsx");

            resultPath = productscombiner.CombineWithProperties(resultPath, @"D:\InstagramExtract\Files\specification_korm.xlsx");
            
            // second step - format products and save into txt file
            var productsFormatter = new ProductsFormatter();
            var txtFilePath = productsFormatter.FormatProductsIntoTxt(resultPath);
            productsFormatter.ConvertToPdf(txtFilePath);

            // third step - extract messages from instagram
            string zipFilePath = @"D:\InstagramExtract\Files\instagram_archive.zip";
            var instagramExtractor = new InstagramMessagingExtract();
            var instagramTxtFilePath = instagramExtractor.ExtractMessages(zipFilePath);

            productsFormatter.ConvertToPdf(instagramTxtFilePath);            
        }
    }
}
