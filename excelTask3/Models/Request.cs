namespace Models
{
    public class Request
    {
        public int Id { get; set; }
        public int ProductId { get; set; }
        public int ClientId { get; set; }
        public int RequestNumber { get; set; }
        public int Amount { get; set; }
        public DateTime Date { get; set; }
    }
}