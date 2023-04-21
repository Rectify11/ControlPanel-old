using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;

namespace Rectify11
{
    public class Program
    {
        public static void Main(string[] cmd)
        {
            try
            {

                Console.WriteLine("Rectify11 control panel extension");
                if(cmd.Length > 0)
                {
                    if (cmd[0].ToLower() == "register")
                    {
                        Console.WriteLine("Registering...");
                        DoRegister();
                    }
                    else if (cmd[0].ToLower() == "unregister")
                    {
                      
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
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            DoRegister();
            Console.WriteLine("press the any key to continue...");
            Console.ReadKey();
        }

        [DllExport]
        public static int DllRegisterServer()
        {
            try
            {
                DoRegister();
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return 10;
            }
        }


        public static void DoRegister()
        {
            //shell:::{0a852434-9b22-36d7-9985-478ccf000690}
            var type = typeof(ThemesPage);
            var assemblyVersion = type.Assembly.GetName().Version.ToString();
            var assemblyFullName = type.Assembly.FullName;
            var className = type.FullName;
            var runtimeVersion = type.Assembly.ImageRuntimeVersion;
            var codeBaseValue = type.Assembly.CodeBase;
            var guid = typeof(ThemesPage).GUID; //must be 0a852434-9b22-36d7-9985-478ccf000690

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
                using (RegistryKey root = clsid.CreateSubKey("{" + guid.ToString() + "}", true))
                {
                    root.SetValue(null, "Rectify11 Settings", RegistryValueKind.String); //display name
                    root.SetValue("Infotip", "Manage your themes and various Rectify11 settings", RegistryValueKind.String); //display name
                    root.SetValue("System.ControlPanel.Category", "5", RegistryValueKind.String);
                    root.SetValue("System.ControlPanel.EnableInSafeMode", 5, RegistryValueKind.DWord);
                    root.SetValue("System.Software.TasksFileUrl", "C:\\tasks.xml", RegistryValueKind.String);
                    //Rectify11.ThemesPage
                    using (RegistryKey progid = root.CreateSubKey("ProgId", true))
                    {
                        progid.SetValue(null, className);
                    }
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
                    using (RegistryKey icon = root.CreateSubKey("DefaultIcon", true))
                    {
                        icon.SetValue(null, Assembly.GetEntryAssembly().Location);
                    }

                    using (RegistryKey shellfolder = root.CreateSubKey("ShellFolder", true))
                    {
                        shellfolder.SetValue("Attributes", (int)(AttributeFlags.IsFolder), RegistryValueKind.DWord);
                    }
                }
            }

            using (RegistryKey progid = Registry.ClassesRoot.CreateSubKey(className, true))
            {
                progid.SetValue(null, className, RegistryValueKind.String); //display name
                using (RegistryKey c = progid.CreateSubKey("CLSID", true))
                {
                    c.SetValue(null, guid, RegistryValueKind.String);
                }
            }

            //Open explorer key
            using (RegistryKey pcNamespace = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\MyComputer\\NameSpace\\", true))
            {
                //create the key
                using (RegistryKey root = pcNamespace.CreateSubKey("{" + guid.ToString() + "}", true))
                {
                    root.SetValue(null, "Rectify11 Settings", RegistryValueKind.String); //display name
                }
            }
            using (RegistryKey pcNamespace = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\ControlPanel\\NameSpace\\", true))
            {
                //create the key
                using (RegistryKey root = pcNamespace.CreateSubKey("{" + guid.ToString() + "}", true))
                {
                    root.SetValue(null, "Rectify11 Settings", RegistryValueKind.String); //display name
                }
            }

            using (RegistryKey pcNamespace = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved\\", true))
            {
                pcNamespace.SetValue("{" + guid.ToString() + "}", "Rectify11 Control Panel Applet");
            }
        }
    }

    [Flags]
    public enum AttributeFlags : uint
    {
        /// <summary>
        /// Indicates that the item can be copied.
        /// </summary>
        CanByCopied = 0x1,

        /// <summary>
        /// Indicates that the item can be moved.
        /// </summary>
        CanBeMoved = 0x2,

        /// <summary>
        /// Indicates that the item can be linked to.
        /// </summary>
        CanBeLinked = 0x4,

        /// <summary>
        /// Indicates that the item is storable and can be bound to the storage system.
        /// </summary>
        IsStorage = 0x00000008,

        /// <summary>
        /// Indicates that the object can be renamed.
        /// </summary>
        CanBeRenamed = 0x00000010,

        /// <summary>
        /// Indicates that the object can be deleted.
        /// </summary>
        CanBeDeleted = 0x00000020,

        /// <summary>
        /// Indicates that the object has property sheets.
        /// </summary>
        HasPropertySheets = 0x00000040,

        /// <summary>
        /// Indicates that the object is a drop target.
        /// </summary>
        IsDropTarget = 0x00000100,

        /// <summary>
        /// Indicates that the object is encrypted and is shown in an alternate colour.
        /// </summary>
        IsEncrypted = 0x00002000,

        /// <summary>
        /// Indicates the object is 'slow' and should be treated by the shell as such.
        /// </summary>
        IsSlow = 0x00004000,

        /// <summary>
        /// Indicates that the icon should be shown with a 'ghosted' icon.
        /// </summary>
        IsGhosted = 0x00008000,

        /// <summary>
        /// Indicates that the item is a link or shortcut.
        /// </summary>
        IsLink = 0x00010000,

        /// <summary>
        /// Indicates that the item is shared.
        /// </summary>
        IsShared = 0x00020000,

        /// <summary>
        /// Indicates that the item is read only.
        /// </summary>
        IsReadOnly = 0x00040000,

        /// <summary>
        /// Indicates that the item is hidden.
        /// </summary>
        IsHidden = 0x00080000,

        /// <summary>
        /// Indicates that item may contain children that are part of the filesystem.
        /// </summary>
        IsFileSystemAncestor = 0x10000000,

        /// <summary>
        /// Indicates that this item is a shell folder, i.e. it can contain 
        /// other shell items.
        /// </summary>
        IsFolder = 0x20000000,

        /// <summary>
        /// Indicates that the object is part of the Windows file system.
        /// </summary>
        IsFileSystem = 0x40000000,

        /// <summary>
        /// Indicates that the item may contain sub folders.
        /// </summary>
        MayContainSubFolders = 0x80000000,

        /// <summary>
        /// Indicates that the object is volatile, and the shell shouldn't cache
        /// data relating to it.
        /// </summary>
        IsVolatile = 0x01000000,

        /// <summary>
        /// Indicates that the object is removable media.
        /// </summary>
        IsRemovableMedia = 0x02000000,

        /// <summary>
        /// Indicates that the object is compressed and should be shown in an alternative colour.
        /// </summary>
        IsCompressed = 0x04000000,

        /// <summary>
        /// The item has children and is browsed with the default explorer UI.
        /// </summary>
        IsBrowsable = 0x08000000,

        /// <summary>
        /// A hint to explorer to show the item bold as it has or is new content.
        /// </summary>
        HasOrIsNewContent = 0x00200000,

        /// <summary>
        /// Indicates that the item is a stream and supports binding to a stream.
        /// </summary>
        IsStream = 0x00400000,

        /// <summary>
        /// Indicates that the item can contain storage items, either streams or files.
        /// </summary>
        IsStorageAncestor = 0x00800000,
    }
}

//hack so that we can build with latest c# version
namespace System.Net.Http { }