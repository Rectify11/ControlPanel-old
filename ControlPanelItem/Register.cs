using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;

namespace Rectify11
{
    public class Program
    {
        public static void Main()
        {
            var type = typeof(ThemesPage);
            var assemblyVersion = type.Assembly.GetName().Version.ToString();
            var assemblyFullName = type.Assembly.FullName;
            var className = type.FullName;
            var runtimeVersion = type.Assembly.ImageRuntimeVersion;
            var codeBaseValue = type.Assembly.CodeBase;
            Console.WriteLine("Rectify11 control panel extension version " + assemblyVersion);
            Console.WriteLine("Registering...");

            //What we need to create:
            //HKEY_CLASSES_ROOT
            //  CLSID
            //    {OUR GUID}
            //      DefaultIcon
            //      InprocServer32
            //        1.0.0.0
            //      ShellFolder

            //Open classes key
            using (RegistryKey clsid = Registry.ClassesRoot.CreateSubKey("CLSID", true))
            {
                //create the key
                var guid = typeof(ThemesPage).GUID; //must be 0a852434-9b22-36d7-9985-478ccf000690
                using (RegistryKey root = clsid.CreateSubKey("{" + guid.ToString() + "}", true))
                {
                    root.SetValue(null, "Rectify11 Settings", RegistryValueKind.String); //display name
                    root.SetValue("Infotip", "Manage your themes and various Rectify11 settings", RegistryValueKind.String); //display name
                    root.SetValue("System.ControlPanel.Category", "5", RegistryValueKind.String);
                    root.SetValue("System.ControlPanel.EnableInSafeMode", 5, RegistryValueKind.DWord);
                    root.SetValue("System.Software.TasksFileUrl", "C:\\tasks.xml", RegistryValueKind.String);

                    using (RegistryKey inprocserver = root.CreateSubKey("InprocServer32", true))
                    {
                        inprocserver.SetValue(null, "mscoree.dll"); //TODO: use .net and not .net framework
                        inprocserver.SetValue("Assembly", assemblyFullName);
                        inprocserver.SetValue("Class", className);
                        inprocserver.SetValue("RuntimeVersion", runtimeVersion);
                        inprocserver.SetValue("ThreadingModel", "Both");
                        inprocserver.SetValue("CodeBase", codeBaseValue);
                        using (RegistryKey versionKey = inprocserver.CreateSubKey(assemblyVersion, true))
                        {
                            versionKey.SetValue("Assembly", assemblyFullName);
                            versionKey.SetValue("Class", className);
                            versionKey.SetValue("RuntimeVersion", runtimeVersion);
                            versionKey.SetValue("CodeBase", codeBaseValue);
                        }
                    }
                }

                using (RegistryKey progid = Registry.ClassesRoot.CreateSubKey(className, true))
                {
                    progid.SetValue(null, className, RegistryValueKind.String); //display name
                    using (RegistryKey c = progid.CreateSubKey("CLSID", true))
                    {
                        progid.SetValue(null, guid, RegistryValueKind.String);
                    }
                }
            }

            //Open explorer key
            using (RegistryKey pcNamespace = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\MyComputer\\NameSpace\\", true))
            {
                //create the key
                var guid = typeof(ThemesPage).GUID; //must be 0a852434-9b22-36d7-9985-478ccf000690
                using (RegistryKey root = pcNamespace.CreateSubKey("{" + guid.ToString() + "}", true))
                {
                    root.SetValue(null, "Rectify11 Settings", RegistryValueKind.String); //display name
                }
            }
        }
    }
}

namespace System.Net.Http { }