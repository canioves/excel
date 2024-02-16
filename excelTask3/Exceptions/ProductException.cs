namespace excelTask3.Exceptions
{
    public class ProductException : NullReferenceException
    {
        public string Value { get; set; }
        public ProductException(string message) : base(message)
        {
        }
        public ProductException(string message, string value) : base(message)
        {
            Value = value;
        }
    }
}