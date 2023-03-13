using ControlPanelItem.COM;
using System.Runtime.InteropServices;

namespace ControlPanelItem
{
    [ComVisible(true)]
    public class ShellItemArray: IShellItemArray
    {
        public ShellItemArray()
        {
        }

        public int BindToHandler([In, MarshalAs(UnmanagedType.Interface)] IntPtr pbc, [In] ref Guid rbhid, [In] ref Guid riid, out IntPtr ppvOut)
        {
            MessageBox.Show("BindToHandler() not implemented");
            throw new NotImplementedException();
        }

        public int GetPropertyStore([In] int Flags, [In] ref Guid riid, out IntPtr ppv)
        {
            MessageBox.Show("GetPropertyStore() not implemented");
            throw new NotImplementedException();
        }

        public int GetPropertyDescriptionList([In] ref PROPERTYKEY keyType, [In] ref Guid riid, out IntPtr ppv)
        {
            MessageBox.Show("GetPropertyDescriptionList() not implemented");
            throw new NotImplementedException();
        }

        int IShellItemArray.GetAttributes(ShellItemAttributeOptions dwAttribFlags, ShellFileGetAttributesOptions sfgaoMask, out ShellFileGetAttributesOptions psfgaoAttribs)
        {
            MessageBox.Show("IShellItemArray.GetAttributes() not implemented");
            throw new NotImplementedException();
        }

        public int GetCount(out uint pdwNumItems)
        {
            MessageBox.Show("GetCount() not implemented");
            throw new NotImplementedException();
        }

        int IShellItemArray.GetItemAt(uint dwIndex, out IShellItem ppsi)
        {
            MessageBox.Show("GetItemAt() not implemented");
            throw new NotImplementedException();
        }

        public int EnumItems([MarshalAs(UnmanagedType.Interface)] out IntPtr ppenumShellItems)
        {
            MessageBox.Show("EnumItems() not implemented");
            throw new NotImplementedException();
        }
    }
}