using ClosedXML.Excel;

namespace excelTask3.Interfaces
{
    public interface IExcelProcess
    {
        public IXLTable GetProductsTable();
        public IXLTable GetClientsTable();
        public IXLTable GetRequestsTable();
    }
}