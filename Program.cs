using System;
using System.Collections.Generic;

// Add nuget reference to smartsheet-csharp-sdk (https://www.nuget.org/packages/smartsheet-csharp-sdk/)
using Smartsheet.Api;
using Smartsheet.Api.Models;
using System.Linq;
using System.Dynamic;
using Newtonsoft.Json.Linq;

namespace sdk_csharp_sample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Initialize client
            SmartsheetClient smartsheet = new SmartsheetBuilder()
                .SetAccessToken("")
                .SetHttpClient(new RetryHttpClient())
                .Build();


            //WebhookResources webhookResources = smartsheet.WebhookResources;
            //Webhook webhook = webhookResources.CreateWebhook(new Webhook
            //{
            //    ApiClientId = "CallbackHandler",
            //    ApiClientName = "S5SmartSheet",
            //    CallbackUrl = "https://abendingar.kopavogur.is/s5smartsheet/Home/CallbackHandler",
            //});

            // List all sheets
            PaginatedResult<Sheet> sheets = smartsheet.SheetResources.ListSheets(new List<SheetInclusion> { SheetInclusion.SHEET_VERSION }, null, null);
            Console.WriteLine("Found " + sheets.TotalCount + " sheets");
            foreach (var sheetsData in sheets.Data) {
                long sheetId = (long)sheetsData.Id;
                Sheet sheet = smartsheet.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);
                Console.WriteLine($"---{sheet.Name} Id:{sheet.Id}");
                if (sheet.Columns != null) {
                    foreach (var column in sheet.Columns) {
                        Console.WriteLine($"{column.Title}  Type:{column.Type}");
                    }
                }
                else {
                    Console.WriteLine("*** sheet.Columns is null");
                }
            }

            if (sheets.TotalCount > 0)
            {
                long sheetId = 844028271978372L;                           // TODO: Uncomment if you wish to read a specific sheet

                Console.WriteLine("Loading sheet id: " + sheetId);

                // Load the entire sheet
                var sheet = smartsheet.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);
                var columnsMap = sheet.Columns.ToDictionary(c => c.Title, c => c.Id);
                Console.WriteLine("Loaded " + sheet.Rows.Count + " rows from sheet: " + sheet.Name);

                // Get a particular row from sheet and dump to Console. This would be how we extract the row data on entry to a webhook callback.
                var rowInclusions = new List<RowInclusion>() { RowInclusion.COLUMN_TYPE, RowInclusion.ATTACHMENTS, RowInclusion.FORMAT, RowInclusion.OBJECT_VALUE, RowInclusion.COLUMN_TYPE};
                var rowExclusions = new List<RowExclusion>();
                Row testRow = smartsheet.SheetResources.RowResources.GetRow(sheetId, 4488531043542916L, null, null);
                dumpRow(testRow, sheet.Columns);

                // Insert a new row.
                var row = new Row() {
                    //ToTop = true, // add the new row to the top of the sheet
                    Cells = new List<Cell>()
                    {
                        new Cell() { ColumnId = columnsMap["Primary Column"], Value = 123 },
                        new Cell() { ColumnId = columnsMap["Text1"], Value = "Einhver texti sem mun endurtaka sig" },
                        new Cell() { ColumnId = columnsMap["OtherText"], Value = "Meiri texti" },
                        new Cell() { ColumnId = columnsMap["Timestamp"], Value = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ") },
                        new Cell() { ColumnId = columnsMap["Column5"], Value = "" },
                        new Cell() { ColumnId = columnsMap["Column6"], Value = "" }
                    }
                };
                var addedRows = smartsheet.SheetResources.RowResources.AddRows(sheetId, new Row[] { row });

                foreach (Row r in addedRows) {
                    dumpRow(r, sheet.Columns);
                }

                // Display the first 5 rows
                foreach (Row r in sheet.Rows.Take(5))
                {
                    dumpRow(r, sheet.Columns);
                }
            }

            Console.WriteLine("Done (Hit enter)");                      // Keep console window open
            Console.ReadLine();
        }

        // Display row contents
        static void dumpRow(Row row, IList<Column> columns)
        {
            Console.WriteLine($"Row#{row.RowNumber} Id#{row.RowNumber}");
            foreach (var cell in row.Cells)
            {
                // Find column name by Id in column collection
                var columName = columns.First(column => column.Id == cell.ColumnId).Title;
                Console.WriteLine("    " + columName + ": " + cell.Value);
            }
        }
    }
}
