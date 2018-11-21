using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace Ch16_P3_ExportDataToOfficeApp
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Car> carsInStock = new List<Car>
            {
                new Car {Color="Green", Make="VW", PetName="Mary"},
                new Car {Color="Red", Make="Saab", PetName="Mel"},
                new Car {Color="Black", Make="Ford", PetName="Hank"},
                new Car {Color="Yellow", Make="BMW", PetName="Davie"}
            };

            ExportToExcel(carsInStock);
            Console.ReadLine();
        }

        static void ExportToExcel(List<Car> carsInStock)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Workbooks.Add();

            Excel._Worksheet worksheet = excelApp.ActiveSheet;

            worksheet.Cells[1, "A"] = "Make";
            worksheet.Cells[1, "B"] = "Color";
            worksheet.Cells[1, "C"] = "Pet Name";

            int row = 1;
            foreach (Car car in carsInStock)
            {
                row++;
                worksheet.Cells[row, "A"] = car.Make;
                worksheet.Cells[row, "B"] = car.Color;
                worksheet.Cells[row, "C"] = car.PetName;
            }

            worksheet.Range["A1"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);

            // Save the file , quit Excel, and display message to user.
            worksheet.SaveAs($@"{Environment.CurrentDirectory}\Inventory.xlsx");
            excelApp.Quit();

            Console.WriteLine(" The Inventory.xlsx file has been saved to your app folder ");
        }

        }
    }
