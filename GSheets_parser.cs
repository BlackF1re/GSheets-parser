using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeOpenXml;
//криво работает покраска, раскраска и фиксирование обновлений

class GSheets_parser
{
    static string credsPath = "externals\\credentials.json";
    static string sheetIdPath = "externals\\spreadsheetId.txt";
    static string range = "ОМС";
    static string parsedPath = "externals\\parsedData.xlsx";
    static int updateCounter = 0;

    static async Task Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        //чтение предшественника
        Dictionary<string, string> oldData = ReadExcelData();

        //авторизация
        var clientSecrets = await GoogleClientSecrets.FromFileAsync(credsPath);
        var credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
            clientSecrets.Secrets,
            new[] { SheetsService.Scope.Spreadsheets },
            "user",
            CancellationToken.None);

        var service = new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential
        });

        string spreadsheetId = File.ReadAllText(sheetIdPath);
        var request = service.Spreadsheets.Values.Get(spreadsheetId, range);
        var response = await request.ExecuteAsync();
        IList<IList<object>> values = response.Values;

        if (values != null && values.Any())
        {
            using (var pkg = new ExcelPackage(new FileInfo(parsedPath)))
            {
                var worksheet = pkg.Workbook.Worksheets["ОМС"] ?? pkg.Workbook.Worksheets.Add("ОМС");

                for (int i = 0; i < values.Count; i++)
                {
                    for (int j = 0; j < values[i].Count; j++)
                    {
                        string value = values[i][j].ToString();

                        //проверка обновлений
                        if (oldData.ContainsKey($"{i + 1},{j + 1}") && oldData[$"{i + 1},{j + 1}"] != value)
                        {
                            worksheet.Cells[i + 1, j + 1].Value = value;

                            // покраска
                            var cell = worksheet.Cells[i + 1, j + 1];
                            var fill = cell.Style.Fill;
                            fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);

                            updateCounter++;

                            Console.WriteLine($"Старое значение ячейки ({i + 1}, {j + 1}): {oldData[$"{i + 1},{j + 1}"]}, новое {value}");
                        }
                        else
                        {
                            //пропуск
                            if (worksheet.Cells[i + 1, j + 1].Value == null)
                                worksheet.Cells[i + 1, j + 1].Value = value;
                        }
                    }
                }

                pkg.Save();
            }

            Console.WriteLine($"Данные обновлены в файле: {parsedPath}");
            Console.WriteLine($"Были обнаружены обновления: {(updateCounter > 0 ? "Да" : "Нет")}");
            if (updateCounter > 0)
            {
                Console.Beep();
                Console.WriteLine($"Количество обновлений: {updateCounter}");
            }
            else
            {
                Console.WriteLine("Обновления отсутствуют.");
            }
        }
        else
        {
            Console.WriteLine("Нет данных");
        }

        Console.ReadKey();
    }

    static Dictionary<string, string> ReadExcelData()
    {
        var data = new Dictionary<string, string>();

        if (File.Exists(parsedPath))
        {
            using (var package = new ExcelPackage(new FileInfo(parsedPath)))
            {
                var worksheet = package.Workbook.Worksheets["ОМС"];
                if (worksheet != null)
                {
                    foreach (var cell in worksheet.Cells)
                    {
                        if (cell.Value != null)
                            data.Add($"{cell.Start.Row},{cell.Start.Column}", cell.Value.ToString());
                    }
                }
            }
        }
        return data;
    }
}
