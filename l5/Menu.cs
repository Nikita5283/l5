using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace l5
{
    public static class Menu
    {
        public static void Start()
        {
            while (true)
            {
                int point = Inspector.CheckInt32("1. Чтение базы данных из excel файла.\n" +
                                                 "2. Просмотр базы данных.\n" +
                                                 "3. Удаление элементов (по ключу).\n" +
                                                 "4. Добавление элементов.\n" +
                                                 "Введите пункт меню: ");
                
                if (point == 1)
                {
                    Read();
                    break;
                }
            }
        }

        public static void Read()
        {
            Workbook wb;
            try
            {
                wb = new Workbook("LR5-var9.xls");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error loading workbook: " + ex.Message);
                return;
            }
            WorksheetCollection wbCollection = wb.Worksheets;
            for (int k = 0; k < wbCollection.Count; k++) 
            {
                Worksheet sheet = wbCollection[k];
                Console.WriteLine(sheet.Name);
                int rows = sheet.Cells.MaxDataRow;
                int cols = sheet.Cells.MaxDataColumn;
                Console.WriteLine($"Кол-во строк: {rows + 1}");
                Console.WriteLine($"Кол-во столбцов: {cols + 1}");

                for (int i = 0; i < rows; i++) 
                { 
                    for (int j = 0; j < cols; j++)
                    {
                        Console.Write(sheet.Cells[i, j].Value + " ");
                        Console.WriteLine();
                    }
                }
            }
        }
    }
}
