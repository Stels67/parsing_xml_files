using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace parsing_xml_files
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(fileName: @"C:\Users\gjghj\OneDrive\Рабочий стол\Exel Files\Test.xlsx");
        }

        static List<PersonModel> GetSetupData()
        {
            List<PersonModel> output = new()
            {
                new() { }
            }
        }
    }
}
