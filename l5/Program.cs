namespace l5
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Workbook wb;
            try
            {
                wb = new Workbook("LR5-var9.xls");
            }
            catch (Exception ex)
            { 
                Console.WriteLine("Error loading workbook: " + ex.Message);
                return;
            }
        }
    }
}
