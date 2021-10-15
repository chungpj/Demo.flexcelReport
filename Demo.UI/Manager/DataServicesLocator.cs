namespace Demo.UI.Manger
{
    public static class DataServicesLocator
    {
        static DataServicesLocator()
        {
            FileManager = new FileManager();
        }
        public static FileManager FileManager { get; private set; }
    }
}