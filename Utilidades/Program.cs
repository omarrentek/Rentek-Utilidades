using System;
using System.Windows.Forms;
using Utilidades.Classes;


namespace Utilidades
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main()
        {
            if (Ini.Run())
            {
                Application.Run();
            }
            else
            {
                System.Windows.Forms.Application.Exit();
            }
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
        }

    }
}
