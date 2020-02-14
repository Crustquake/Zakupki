namespace Zakupki
{
    public class Program
    {
        static void Main(string[] args)
        {
            using (OrdersProcessing ordersProcessing = new OrdersProcessing())
            {
                ordersProcessing.Run();
            }
        }
    }
}
