using excelTask3.Interfaces;
using Models;

namespace excelTask3.Service
{
    public class ProductService : ModelService
    {
        public List<Product> ProductsList { get { return AllProducts; } }
        public ProductService(IExcelProcess excelProcess) : base(excelProcess)
        {
        }
        public Product GetProductByName(string name)
        {
            var products = ProductsList;
            var result = products.SingleOrDefault(x => x.Name == name);
            return result;
        }

        public Product GetProductById(int id)
        {
            var products = ProductsList;
            var result = products.SingleOrDefault(x => x.Id == id);
            return result;
        }
    }
}