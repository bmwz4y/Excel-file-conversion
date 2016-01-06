using System;
using System.Windows.Forms;

namespace test
{
    public static class Program
    {
        public static string strPath = Environment.GetFolderPath(
                         System.Environment.SpecialFolder.DesktopDirectory);
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
