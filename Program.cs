namespace TestApp
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            var app = new TestApplication();
            await app.Run();
        }
    }
}
