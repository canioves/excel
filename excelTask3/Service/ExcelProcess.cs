using ClosedXML.Excel;
using excelTask3.Interfaces;

namespace excelTask3.Service
{
    public class ExcelProcess : IExcelProcess
    {
        private IXLWorksheet Products;
        private IXLWorksheet Clients;
        private IXLWorksheet Requests;

        public ExcelProcess(string path)
        {
            string dirPath = Directory.GetCurrentDirectory();
            string fullPath = dirPath.Replace("bin\\Debug\\net8.0", "") + path;
            XLWorkbook workbook = new XLWorkbook(fullPath);
            IXLWorksheets worksheets = workbook.Worksheets;
            Console.WriteLine($"Успешно открыт файл по пути: {fullPath}");

            Products = worksheets.Worksheet(1);
            Clients = worksheets.Worksheet(2);
            Requests = worksheets.Worksheet(3);

        }

        public IXLTable GetProductsTable() => Products.RangeUsed().CreateTable();
        public IXLTable GetClientsTable() => Clients.RangeUsed().CreateTable();
        public IXLTable GetRequestsTable() => Requests.RangeUsed().CreateTable();
    }
}