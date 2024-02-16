using ClosedXML.Excel;
using DocumentFormat.OpenXml.Presentation;
using excelTask3.Interfaces;

namespace excelTask3.Service
{
    public class ExcelProcess : IExcelProcess
    {
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
                XLWorkbook workbook = new XLWorkbook(fullPath);
                IXLWorksheets worksheets = workbook.Worksheets;
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
    }
}