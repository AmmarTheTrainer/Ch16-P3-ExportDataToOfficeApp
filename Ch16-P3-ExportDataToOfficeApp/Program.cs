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

        }

        static void ExportToExcel(List<Car> carsInStock)
        {
            //
        }
        }
    }
