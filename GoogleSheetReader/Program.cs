using System;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace GoogleSheetReader
{
    static class Program
    {
        private const string ApplicationName = "Cautionary Contacts Google Sheet Reader";
        private static readonly string SpreadsheetId = Environment.GetEnvironmentVariable("SPREADSHEET_ID");

        private static void Main()
        {
            var credentials = GoogleCredential
                .FromFile("../../../key.json")
                .CreateScoped(SheetsService.Scope.SpreadsheetsReadonly);

            var service = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credentials,
                ApplicationName = ApplicationName
            });

            var request = service.Spreadsheets.Get(SpreadsheetId);
            request.IncludeGridData = true;
            var response = request.Execute();

            var rows = response
                .Sheets.First()
                .Data.First()
                .RowData.Select(
                    rd => rd.Values.Where(
                        cd => cd.FormattedValue is not null).Select(
                        cd => cd.FormattedValue.ToString()
                    ).ToList()
                ).ToList();

            var headers = rows.First();
            if (ValidateHeaders<ExampleModel>(headers) is false)
                throw new Exception("Spreadsheet schema does not match the provided type!");

            PrintSpreadSheet(rows);

            rows.RemoveAt(0);

            var highlightedFlags = response
                .Sheets.First()
                .Data.First()
                .RowData.Select(
                    rd => rd.Values.All(
                        cd => cd.EffectiveFormat.BackgroundColor.Red.Equals(.800000012f)
                    ).ToString()
                );

            var pRows = rows
                .Pivot()
                .Append(highlightedFlags)
                .Pivot();

            var entities = pRows.Select(row => GetEntityFromRow(row.ToList()));

            Console.Write(JsonConvert.SerializeObject(entities, Formatting.Indented, new StringEnumConverter()) +
                          Environment.NewLine);
        }

        private static bool ValidateHeaders<T>(List<string> headers)
        {
            // TODO: Reflect the provided class and compare against header strings
            // TODO: Ignore boolean class properties - they weren't a column in the spreadsheet,
            // TODO: and so won't be reflected in the headers List
            return true;
        }

        private static ExampleModel GetEntityFromRow(List<string> row)
        {
            return new()
            {
                Name = row[0],
                Number = int.Parse(row[1]),
                IsHighlighted = bool.Parse(row[2])
            };
        }

        private static void PrintSpreadSheet(List<List<string>> rows)
        {
            var padding = (
                from row in rows
                from string cell in row
                select cell.Length
            ).Prepend(0).Max();

            foreach (var row in rows.ToArray())
            {
                Console.Write("| ");
                foreach (var cell in row) Console.Write($"{cell.PadRight(padding)} | ");
                Console.WriteLine();
            }
        }

        private static IEnumerable<IEnumerable<T>> Pivot<T>(this IEnumerable<IEnumerable<T>> source)
        {
            var enumerators = source.Select(e => e.GetEnumerator()).ToArray();

            try
            {
                while (enumerators.All(e => e.MoveNext()))
                    yield return enumerators
                        .Select(e => e.Current)
                        .ToArray();
            }
            finally
            {
                Array.ForEach(enumerators, e => e.Dispose());
            }
        }
    }
}