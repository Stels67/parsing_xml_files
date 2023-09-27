using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;

namespace parsing_xml_files
{
    class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(@"C:\Users\gjghj\OneDrive\Рабочий стол\Exel Files\Test.xlsx");

            var people = GetSetupData();

            await SaveExcelFile(people, file);

            List<PersonModel> peopleFromExcel = await LoadExcelFile(file);

            foreach (var p in peopleFromExcel)
            {
                Console.WriteLine($"{p.FirstName} {p.Surname} {p.PhoneNumber} {p.Email} {p.PassportNumber}");
            }
        }

        private static async Task<List<PersonModel>> LoadExcelFile(FileInfo file)
        {
            List<PersonModel> output = new();

            using var package = new ExcelPackage(file);

            await package.LoadAsync(file);

            var ws = package.Workbook.Worksheets[0];

            int row = 3;
            int col =1;

            while (string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()) == false)
            {
                PersonModel p = new();
                p.FirstName = ws.Cells[row, col].Value.ToString();
                p.Surname = ws.Cells[row, col+1].Value.ToString();
                p.PhoneNumber = ws.Cells[row, col+2].Value.ToString();
                p.Email = ws.Cells[row, col+3].Value.ToString();
                p.PassportNumber = ws.Cells[row, col + 4].Value.ToString();
                output.Add(p);
                row += 1;
            }

            return output;

        }

        private static async Task SaveExcelFile(List<PersonModel> people, FileInfo file)
        {
            DeleteIfExists(file);

            using var package = new ExcelPackage(file);

            var ws = package.Workbook.Worksheets.Add("MainReport");

            var range = ws.Cells["A2"].LoadFromCollection(people, true);
            range.AutoFitColumns();

            ws.Cells["A1"].Value = "Список Людей";
            ws.Cells["A1:E1"].Merge = true;
            ws.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Row(1).Style.Font.Size = 24;
            ws.Row(1).Style.Font.Color.SetColor(Color.Blue);

            ws.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Row(2).Style.Font.Bold = true;
            ws.Column(3).Width = 20;

            await package.SaveAsync();
        }

        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }

        private static List<PersonModel> GetSetupData()
        {
            List<PersonModel> output = new()
            {
                new() { FirstName = "Василий", Surname = "Васильев", PhoneNumber = "+79254898547", Email = "майл пользователя", PassportNumber = "45484547512" },
                new() { FirstName = "Василий", Surname = "Васильева", PhoneNumber = "+7925489854", Email = "майл пользователя", PassportNumber = "45484547512" },
                new() { FirstName = "Василий", Surname = "Васильеву", PhoneNumber = "+792548985", Email = "майл пользователя", PassportNumber = "45484547512" }
            };

            return output;
        }
    }
}
