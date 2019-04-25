using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Net;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using Microsoft.Azure.Storage;
using Microsoft.Azure.Storage.Blob;

namespace FunctionApp1
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ExecutionContext context,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            //// 1. Get data from SQL

            DataSet ds = new DataSet();

            using (SqlConnection connection = new SqlConnection(Environment.GetEnvironmentVariable("sqlConnection")))
            {
                using (SqlDataAdapter dataAdapter = new SqlDataAdapter("select * from Test", connection))
                {
                    dataAdapter.Fill(ds);
                }
            }

            //// 2. Create and Format the excel

            using (SpreadsheetDocument workbook = SpreadsheetDocument.Create(context.FunctionDirectory + "\\test.xlsx", SpreadsheetDocumentType.Workbook))
            {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookPart = workbook.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Test Sheet" };

                sheets.Append(sheet);
                workbookPart.Workbook.Save();

                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                Row row = new Row();

                foreach (DataTable dataTable in ds.Tables)
                {
                    foreach (DataRow dataRow in dataTable.Rows)
                    {
                        row = new Row();
                        row.Append(
                            ConstructCell(dataRow["id"].ToString(), CellValues.String),
                            ConstructCell(dataRow["name"].ToString(), CellValues.String));

                        sheetData.AppendChild(row);
                    }
                }

                workbookPart.Workbook.Save();
            }

            //// 3. Uplaod the resulting excel to Storage blob
            
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(Environment.GetEnvironmentVariable("MyStorageConnectionAppSetting"));
            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();
            CloudBlobContainer container = blobClient.GetContainerReference("test");
            CloudBlockBlob blockBlob = container.GetBlockBlobReference("test.xlsx");

            blockBlob.UploadFromFile(context.FunctionDirectory + "\\test.xlsx");

            return new OkResult();
        }

        static Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }
    }
}
