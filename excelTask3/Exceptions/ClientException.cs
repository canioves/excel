namespace excelTask3.Exceptions
{
    public class ClientException : NullReferenceException
    {
        public int Value { get; set; }
        public ClientException(string message) : base(message)
        {
        }

        public ClientException(string message, int value) : base(message)
        {
            Value = value;
        }
    }
}