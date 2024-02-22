using Amazon;
using Amazon.S3;
using Amazon.S3.Model;
using Amazon.Textract;
using Amazon.Textract.Model;
using Aspose.Pdf.Devices;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;

namespace AWSTextract
{
    class Program
    {

        /*
         * In this example, we will read a PDF INVOICE and export its data.
         * We optionally convert the PDF to an image to skip uploading the document to an s3 bucket.
         * 
        */


        //Path of our local PDF file
        static readonly string file = @"C:\Users\example\Downloads\Invoice.pdf";
        //Path where we will save the results
        static readonly string resultsFolder = @"C:\Users\example\Downloads\";
        //Amazon settings
        static readonly string s3Bucket = "example-bucket-name";
        static readonly AmazonS3Config config = new() { ProxyCredentials = new NetworkCredential("user", "key"), RegionEndpoint = RegionEndpoint.USEast1 };
        static readonly AmazonTextractClient textractClient = new(RegionEndpoint.USEast1);


        static async Task Main()
        {
            //await AnalyzePDF();
            await AnalyzeIMG();
        }

        static async Task AnalyzePDF()
        {
            try
            {
                AmazonS3Client s3Client = new(config);

                PutObjectRequest putRequest = new()
                {
                    BucketName = s3Bucket,
                    FilePath = file,
                    Key = Path.GetFileName(file)
                };
                await s3Client.PutObjectAsync(putRequest);

                StartExpenseAnalysisResponse startResponse = await textractClient.StartExpenseAnalysisAsync(new StartExpenseAnalysisRequest()
                {
                    DocumentLocation = new()
                    {
                        S3Object = new()
                        {
                            Bucket = s3Bucket,
                            Name = putRequest.Key
                        }
                    },
                });
                GetExpenseAnalysisRequest getAnalysisRequest = new()
                {
                    JobId = startResponse.JobId
                };

                GetExpenseAnalysisResponse getAnalysisResponse = null;
                do
                {
                    Thread.Sleep(1000);
                    getAnalysisResponse = await textractClient.GetExpenseAnalysisAsync(getAnalysisRequest);
                } while (getAnalysisResponse.JobStatus == JobStatus.IN_PROGRESS);

                if (getAnalysisResponse.JobStatus == JobStatus.SUCCEEDED)
                {
                    getAnalysisResponse.ExpenseDocuments.ForEach(document =>
                    {
                        ExportJson(JsonConvert.SerializeObject(document));
                        ExportXLSX(document);
                    });
                }
                else
                {
                    Console.WriteLine($"Process failed with message: {getAnalysisResponse.StatusMessage}");
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("An exception occurred: " + ex.Message);
            }
        }
        static async Task AnalyzeIMG()
        {
            //Convert the PDF to a PNG image
            Aspose.Pdf.Document document = new(file);
            PngDevice renderer = new();
            MemoryStream ms = new();
            renderer.Process(document.Pages[1], ms);

            try
            {
                AmazonS3Client s3Client = new(config);

                AnalyzeExpenseResponse response = await textractClient.AnalyzeExpenseAsync(new AnalyzeExpenseRequest()
                {
                    Document = new()
                    {
                        Bytes = ms
                    }
                });

                if (response.HttpStatusCode == HttpStatusCode.OK)
                {
                    response.ExpenseDocuments.ForEach(document =>
                    {
                        ExportJson(JsonConvert.SerializeObject(document));
                        ExportXLSX(document);
                    });
                }
                else
                {
                    Console.WriteLine($"Process failed with message: {response.HttpStatusCode}");
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("An exception occurred: " + ex.Message);
            }
        }

        static void ExportJson(string json)
        {
            StreamWriter sw = new(@$"{resultsFolder}\InvoiceAnalysisResult.json");
            sw.WriteLine(json);
            sw.Close();
        }
        static void ExportXLSX(ExpenseDocument document)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelTextFormat format = new() { Delimiter = '!' };

            using (ExcelPackage package = new())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add($"{document.ExpenseIndex}");
                worksheet.DefaultColWidth = 30;

                //We can read specific values ​​from an invoice by passing the field id or the label text, which we can get by parsing the exported json.

                ExpenseField invoiceID = document.SummaryFields.Find(x => x.Type.Text == "INVOICE_RECEIPT_ID") ?? null;
                if (invoiceID == null)
                {
                    invoiceID = document.SummaryFields.Find(x => x.LabelDetection?.Text == "INVOICE") ?? null;
                }
                worksheet.Cells["A1"].Value = "INVOICE";
                worksheet.Cells["A2"].Value = invoiceID?.ValueDetection.Text ?? "";

                ExpenseField date = document.SummaryFields.Find(x => x.Type.Text == "INVOICE_RECEIPT_DATE") ?? null;
                if (date == null)
                {
                    date = document.SummaryFields.Find(x => x.LabelDetection?.Text == "Date Issued:");
                }
                worksheet.Cells["B1"].Value = "DATE";
                worksheet.Cells["B2"].Value = date?.ValueDetection.Text ?? "";
                worksheet.Cells["B2"].Style.Numberformat.Format = "dd/mm/yyyy hh:mm";

                ExpenseField total = document.SummaryFields.Find(x => x.Type.Text == "TOTAL") ?? null;
                if (total == null)
                {
                    total = document.SummaryFields.Find(x => x.LabelDetection?.Text == "Amount Due");
                }
                worksheet.Cells["C1"].Value = "TOTAL";
                worksheet.Cells["C2"].Value = total?.ValueDetection.Text ?? "";

                #region Style 
                ExcelRangeBase encabezadoHeader = worksheet.Cells["A1:C1"];
                encabezadoHeader.Style.Fill.PatternType = ExcelFillStyle.Solid;
                encabezadoHeader.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(220, 214, 218));
                encabezadoHeader.Style.Font.Bold = true;
                encabezadoHeader.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                encabezadoHeader.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                encabezadoHeader.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                encabezadoHeader.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                encabezadoHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ExcelRangeBase encabezadoBody = worksheet.Cells["A2:C2"];
                encabezadoBody.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                encabezadoBody.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                encabezadoBody.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                encabezadoBody.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                encabezadoBody.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                #endregion

                //To access the invoice items we can iterate over them
                document.LineItemGroups.ForEach(lineItemGroup =>
                {
                    string header = "";
                    int totalCols = 0;
                    int totalRows = lineItemGroup.LineItems.Count;

                    lineItemGroup.LineItems.FirstOrDefault().LineItemExpenseFields.ForEach(lineItemExpenseField =>
                    {
                        if (lineItemExpenseField.LabelDetection != null && lineItemExpenseField.Type.Text != "EXPENSE_ROW")
                        {
                            header += lineItemExpenseField.LabelDetection.Text + "!";
                            totalCols++;
                        }
                    });
                    header = header[0..^1];
                    ExcelRangeBase tableHeader = worksheet.Cells[5, 1, 5, totalCols].LoadFromText(header, format);
                    #region Style
                    tableHeader.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    tableHeader.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(220, 214, 218));
                    tableHeader.Style.Font.Bold = true;
                    tableHeader.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    tableHeader.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    tableHeader.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    tableHeader.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    tableHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    tableHeader.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    tableHeader.Style.WrapText = true;
                    #endregion

                    int row = 6;
                    lineItemGroup.LineItems.ForEach(lineItem =>
                    {
                        string item = "";
                        lineItem.LineItemExpenseFields.ForEach(lineItemExpenseField =>
                        {
                            if (lineItemExpenseField.LabelDetection != null && lineItemExpenseField.Type.Text != "EXPENSE_ROW")
                            {
                                item += lineItemExpenseField.ValueDetection.Text + "!";
                            }
                        });
                        item = item[0..^1];

                        ExcelRangeBase tableRow = worksheet.Cells[row, 1, row, totalCols].LoadFromText(item, format);
                        #region Style
                        tableRow.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        tableRow.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        tableRow.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        tableRow.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        tableRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        tableRow.Style.WrapText = true;
                        #endregion
                        row++;
                    });
                });

                worksheet.Name = invoiceID?.ValueDetection.Text + " " + document.ExpenseIndex.ToString() ?? document.ExpenseIndex.ToString();
                package.SaveAs(resultsFolder + worksheet.Name + ".xlsx");
            };
        }

    }
}