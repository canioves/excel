using excelTask3.Interfaces;
using Models;

namespace excelTask3.Service
{
    public class ProductService : ModelService
    {
        public ProductService(IExcelProcess excelProcess) : base(excelProcess)
        {

        }
        public List<Product> GetAllProducts()
        {
            var table = _excelProcess.GetProductsTable();
            var rowsCount = table.RowCount();
            List<Product> products = new List<Product>();
            for (int i = 2; i < rowsCount; i++)
            {
                var product = new Product
                {
                    Id = Convert.ToInt32(table.Field("Код товара").DataCells.Single(x => x.Address.RowNumber == i).Value.ToString()),
                    Name = table.Field("Наименование").DataCells.Single(x => x.Address.RowNumber == i).Value.ToString(),
                    UnitName = table.Field("Ед. измерения").DataCells.Single(x => x.Address.RowNumber == i).Value.ToString(),
                    Price = Convert.ToDouble(table.Field("Цена товара за единицу").DataCells.Single(x => x.Address.RowNumber == i).Value.ToString())
                };
                products.Add(product);
            }
            return products;

        }
        public Product GetProductByName(string name)
        {
            var products = GetAllProducts();
            var result = products.Single(x => x.Name == name);
            return result;
        }
    }
}