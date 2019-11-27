using System;
using ExcelDataReaderLib;

namespace ExcelApp
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelDataReader xlLib = new ExcelDataReader();
            Console.WriteLine("Enter the name You want to search");
            String Name = Console.ReadLine();
            string result = xlLib.GetExcelData(Name);
            Console.WriteLine(result);
            Console.ReadLine();
        }
    }
}
