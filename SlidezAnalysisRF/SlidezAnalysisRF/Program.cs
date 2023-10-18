namespace SlidezAnalysisRF
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {

            //Aspose.Slides.License license = new Aspose.Slides.License();


            //license.SetLicense(@"c:\rlz\aspose\license.txt");
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }
    }
}