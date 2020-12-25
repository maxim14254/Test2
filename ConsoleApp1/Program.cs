using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Write("Введите путь к файлу с данными о насосах в виде Excel: ");
            string pathData = Console.ReadLine();//переменная, которая хранит путь к файлу с данными о насосах
            Console.Write("Введите путь к файлу сопоставления в виде Excel: ");
            string pathMapping = Console.ReadLine();//переменная, которая хранит путь к файлу сопоставления
            Console.Write("Введите название экспортного файла: ");
            string exportFile = Console.ReadLine();// переменная, которая хранит название выходного файла
            List <Pump> pumps = new List<Pump>();// коллекция насосов, которые будем сериализовать
            DataContractJsonSerializer json = new DataContractJsonSerializer(typeof(List<Pump>));// экземпляр класса сериализации

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;//лицензия (нужно для работы пакета EPPlus)

                ExcelPackage excelFile = new ExcelPackage(new FileInfo(pathData));// эклемпляр класса, в котором находится файл с данными о насосах 

                ExcelWorksheet worksheet = excelFile.Workbook.Worksheets[0];// экземпляр класса в котором находится первый лист из файла с данными о насосах 

                excelFile = new ExcelPackage(new FileInfo(pathMapping)); // эклемпляр класса, в котором находится файл сопоставления

                ExcelWorksheet SubType = excelFile.Workbook.Worksheets.Where(e => e.Name.ToLower() == "subtype").First();// экземпляр класса в котором находится лист файла сопоставления с названием SubType (регистр не имеет значения)
                ExcelWorksheet Type = excelFile.Workbook.Worksheets.Where(e => e.Name.ToLower() == "type").First();// экземпляр класса в котором находится лист файла сопоставления с названием Type (регистр не имеет значения)

                for (int i = 2; i <= worksheet.Dimension.End.Row; i++)// перебирам файл с данными о насосах (первую строку не трогаем)
                {
                    int rowSub = SubType.Cells[1, 1, SubType.Dimension.End.Row, 1].Where(t => t.Text.ToLower() == worksheet.Cells[i, 3].Text.ToLower()).First().End.Row;// в файле сопоставления на листе "SubType" получаем номер строки, которая удовлетворяет условию (файл с данными).SubType == (файл сопоставления).SubType
                    int rowType = Type.Cells[1, 1, Type.Dimension.End.Row, 1].Where(t => t.Text.ToLower() == worksheet.Cells[i, 2].Text.ToLower()).First().End.Row;// в файле сопоставления на листе "Type" получаем номер строки, которая удовлетворяет условию (файл с данными).Type == (файл сопоставления).Type

                    pumps.Add(new Pump// заносим в коллекцию новый насос
                    {
                        Code = worksheet.Cells[i, 1].Text,// Code берем из файла с данными
                        D = Convert.ToInt32(double.Parse(worksheet.Cells[i, 5].Text) * 1000),// D берем из файла с данными и умножаем на 1000
                        Fluid = GetFluid(worksheet.Cells[i, 6].Text),// Fluid берем из файла с данными
                        SafetyClass = GetSafetyClass(worksheet.Cells[i, 4].Text),// SafetyClass берем из файла с данными
                        Weight = double.Parse(worksheet.Cells[i, 7].Text) / 1000,// Weight берем из файла с данными и делим на 1000
                        Type = Type.Cells[rowType, 2].Text,// Type берем из файла сопоставления
                        SubType = SubType.Cells[rowSub, 2].Text,// SubType берем из файла сопоставления
                    });
                }
                using (var file = new FileStream($"{exportFile}.json", FileMode.Create))// создаем файл json
                {
                    json.WriteObject(file, pumps);// сериализация
                    Console.Write($"Готово. Файл {exportFile}.json находится в папке Bin");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            Console.ReadKey();
        }
        private static int GetSafetyClass(string safetyClass)// класс для перевода римские цифры в арабские до 39
        {
            char[] charSafetyClass = safetyClass.ToCharArray();// разбиваем на символы полученную строку
           
            Dictionary<char, int> numbres = new Dictionary<char, int>(3);//коллекция (ключ соответствует арабскому значению значению)
            numbres.Add('I', 1);
            numbres.Add('V', 5);
            numbres.Add('X', 10);

            int[] sum = new int[charSafetyClass.Length];
            for (int i = 0; i < charSafetyClass.Length; i++) // перебираем массив символов
            {
                sum[i] = numbres.Where(n => n.Key == charSafetyClass[i]).First().Value;// записываем в массив результат, который удовлетворяем условию n.Key == charSafetyClass[i]
                if (i > 0 && sum[i] > sum[i - 1])// если последующая цифра больше предыдущей, вычитаем 2 (это либо 9 либо 4)
                {
                    sum[i] = sum[i] - 2;
                }
            }

            return sum.Sum();// возвращаем сумму полученных цифр
        }
        private static string GetFluid(string safetyClass)// класс для преобразования данных Fluid
        {
            string res = null;
            if(safetyClass== "water")
            {
                res = "вода";
            }
            else if(safetyClass == "acid")
            {
                res = "кислота";
            }
            return res;
        }
    }
}
