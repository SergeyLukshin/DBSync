﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace DBSync
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]

        static void Main()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new DBSync());
            }
            catch (Exception ex)
            {
            }
        }
    }
}