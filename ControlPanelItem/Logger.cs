using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ControlPanelItem
{
    public static class Logger
    {
        public static string filePath = @"C:\Users\Misha\shellog.txt";
        public static void Log(string message)
        {
            Console.WriteLine(message);
            Debug.WriteLine(message);
            using (StreamWriter streamWriter = File.AppendText(filePath))
            {
                streamWriter.WriteLine(message);
                streamWriter.Close();
            }
        }
    }
}
