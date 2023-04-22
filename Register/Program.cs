using Rectify11;

namespace Register
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {

                Console.WriteLine("Rectify11 control panel extension with argument length of "+args.Length);
                if (args.Length > 0)
                {
                    if (args[0].ToLower() == "register")
                    {
                        Console.WriteLine("Registering...");
                        RegisterUtil.DoRegister();
                    }
                    else if (args[0].ToLower() == "unregister")
                    {
                        Console.WriteLine("Uninstalling...");
                        RegisterUtil.DoUnregister();
                    }
                    else
                    {
                        Console.WriteLine("Usage: rectify11 register / unregister");
                    }
                }
                else
                {
                    Console.WriteLine("Usage: rectify11 register / unregister");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            Console.WriteLine("press the any key to continue...");
            Console.ReadKey();
        }
    }
}