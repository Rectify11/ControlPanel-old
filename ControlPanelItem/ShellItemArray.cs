using Rectify11.COM;
using System.Runtime.InteropServices;

namespace Rectify11
{
    [ComVisible(true)]
    [Guid("A0360902-29E5-4AE1-851E-8EDAB9007D68")]
    public class ShellItemArray: IShellItemArray
    {
        public ShellItemArray()
        {
        }

        public int BindToHandler([In, MarshalAs(UnmanagedType.Interface)] IntPtr pbc, [In] ref Guid rbhid, [In] ref Guid riid, out IntPtr ppvOut)
        {
            Logger.Log("BindToHandler() not implemented");
            throw new NotImplementedException();
        }

        public int GetPropertyStore([In] int Flags, [In] ref Guid riid, out IntPtr ppv)
        {
            Logger.Log("GetPropertyStore() not implemented");
            throw new NotImplementedException();
        }

        public int GetPropertyDescriptionList([In] ref PROPERTYKEY keyType, [In] ref Guid riid, out IntPtr ppv)
        {
            Logger.Log("GetPropertyDescriptionList() not implemented");
            throw new NotImplementedException();
        }

        int IShellItemArray.GetAttributes(ShellItemAttributeOptions dwAttribFlags, ShellFileGetAttributesOptions sfgaoMask, out ShellFileGetAttributesOptions psfgaoAttribs)
        {
            Logger.Log("IShellItemArray.GetAttributes() not implemented");
            throw new NotImplementedException();
        }

        public int GetCount(out uint pdwNumItems)
        {
            Logger.Log("GetCount() not implemented");
            throw new NotImplementedException();
        }

        int IShellItemArray.GetItemAt(uint dwIndex, out IShellItem ppsi)
        {
            Logger.Log("GetItemAt() not implemented");
            throw new NotImplementedException();
        }

        public int EnumItems([MarshalAs(UnmanagedType.Interface)] out IntPtr ppenumShellItems)
        {
            Logger.Log("EnumItems() not implemented");
            throw new NotImplementedException();
        }
    }
}