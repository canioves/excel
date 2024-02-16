using excelTask3.Service;
using Models;

Start start = new Start();

string path = @"Data\dataTable.xlsx";
start.Init(path);
// start.PrintAllProducts();
start.PrintProductId("Йогур");


public class Start
{
    private ProductService productService;
    private ExcelProcess excelProcess;
    public void Init(string path)
    {
        excelProcess = new ExcelProcess(path);
        productService = new ProductService(excelProcess);
    }
    public void PrintAllProducts()
    {
        Console.WriteLine("Все продукты:");
        var lst = productService.GetAllProducts();
        foreach (var product in lst)
        {
            Console.WriteLine($"Айди: {product.Id}");
            Console.WriteLine($"Название: {product.Name}");
            Console.WriteLine($"Единица измерения: {product.UnitName}");
            Console.WriteLine($"Цена: {product.Price}\n");
        }
    }
    public int PrintProductId(string name)
    {
        Product product = productService.GetProductByName(name);
        return product == null ? 0 : product.Id;
    }
}