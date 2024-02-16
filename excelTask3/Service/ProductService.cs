using excelTask3.Interfaces;
using Models;

namespace excelTask3.Service
{
    public class ProductService : ModelService
    {
        public ProductService(IExcelProcess excelProcess) : base(excelProcess)
        {
        }
        public Product GetProductByName(string name)
        {
            var products = AllProducts;
            var result = products.Single(x => x.Name == name);
            return result;
        }

        public Product GetProductById(int id)
        {
            var products = AllProducts;
            var result = products.Single(x => x.Id == id);
            return result;
        }
    }
}