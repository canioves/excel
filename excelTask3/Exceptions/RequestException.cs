namespace excelTask3.Exceptions
{
    public class RequestException : NullReferenceException
    {
        public string Value { get; set; }
        public RequestException(string message) : base(message)
        {
        }
        public RequestException(string message, string value) : base(message)
        {
            Value = value;
        }
    }
}