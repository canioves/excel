using System.Data;
using ClosedXML.Excel;
using excelTask3.Interfaces;
using Models;

namespace excelTask3.Service
{
    public class ExcelProcess : IExcelProcess
    {
        private IXLWorkbook Workbook;
        private IXLWorksheet Products;
        private IXLWorksheet Clients;
        private IXLWorksheet Requests;
        public bool repeatInit = true;
        public ExcelProcess(string path)
        {
            string dirPath = Directory.GetCurrentDirectory();
            string fullPath = dirPath.Replace("bin\\Debug\\net8.0", "") + path;
            try
            {
                Workbook = new XLWorkbook(fullPath);
                IXLWorksheets worksheets = Workbook.Worksheets;
                Products = worksheets.Worksheet(1);
                Clients = worksheets.Worksheet(2);
                Requests = worksheets.Worksheet(3);
                

                Console.WriteLine($"Успешно открыт файл по пути: {fullPath}\n");
                repeatInit = false;
            }
            catch
            {
                Console.WriteLine($"Ошибка открытия файла по пути: {fullPath}. Попробуйте еще раз.\n");
                repeatInit = true;
            }
        }
        private IXLTable CreateTables(IXLWorksheet sheet)
        {
            var firstCell = sheet.FirstCellUsed();
            var lastCell = sheet.LastCellUsed();
            var range = sheet.Range(firstCell.Address, lastCell.Address);
            return range.AsTable();
        }
        
        public IXLTable GetProductsTable() => CreateTables(Products);
        public IXLTable GetClientsTable() => CreateTables(Clients);
        public IXLTable GetRequestsTable() => CreateTables(Requests);

        public void UpdateClientsTable(IXLTable clientsTable)
        {
            var firstCell = Clients.FirstCellUsed().CellBelow().Address;
            var lastCell = Clients.LastCellUsed().Address;

            for (int i = firstCell.RowNumber; i < lastCell.RowNumber; i++)
            {
                for (int j = firstCell.ColumnNumber; j < lastCell.ColumnNumber; j++)
                {
                    Clients.Cell(i, j).Value = clientsTable.Cell(i, j).Value;
                }
            }
            Workbook.Save();
        }
    }
}