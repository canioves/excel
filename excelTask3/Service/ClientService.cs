using excelTask3.Interfaces;

namespace excelTask3.Service
{
    public class ClientService : ModelService
    {
        public ClientService(IExcelProcess excelProcess) : base(excelProcess)
        {

        }

        // public Product GetProductByName(string name)
        // {
        //     var product = new Product();
        //     var table = _excelProcess.GetProductsTable();
        //     int productRowId = table.Field("Наименование").DataCells.Single(x => x.Value.ToString() == name).Address.RowNumber;

        //     product.Id = Convert.ToInt32(table.Field("Код товара").DataCells.Single(x => x.Address.RowNumber == productRowId).Value.ToString());
        //     product.Name = table.Field("Код товара").DataCells.Single(x => x.Address.RowNumber == productRowId).Value.ToString();
        //     product.UnitName = table.Field("Код товара").DataCells.Single(x => x.Address.RowNumber == productRowId).Value.ToString();
        //     product.Price = Convert.ToDouble(table.Field("Цена товара за единицу").DataCells.Single(x => x.Address.RowNumber == productRowId).Value.ToString());

        //     return product;
        // }
    }
}