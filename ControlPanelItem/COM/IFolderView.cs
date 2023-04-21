using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Rectify11.COM
{
    [ComImport,
       Guid("cde725b0-ccc9-4519-917e-325d72fab4ce"),
       InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IFolderView
    {
        void GetCurrentViewMode([Out] out uint pViewMode);

        int SetCurrentViewMode(uint ViewMode);
        unsafe int GetFolder(ref Guid riid, [Out, MarshalAs(UnmanagedType.IUnknown)] out object ppv);
        void Item(int iItemIndex, out IntPtr ppidl);
        void ItemCount(uint uFlags, out int pcItems);
        void Items(uint uFlags, ref Guid riid, [Out, MarshalAs(UnmanagedType.IUnknown)] out object ppv);
        void GetSelectionMarkedItem(out int piItem);
        void GetFocusedItem(out int piItem);
        int GetItemPosition(IntPtr pidl, [MarshalAs(UnmanagedType.LPStruct)] out POINT ppt);
        void GetSpacing([Out] out POINT ppt);
        void GetDefaultSpacing(out POINT ppt);
        void GetAutoArrange();
        void SelectItem(int iItem, uint dwFlags);
        void SelectAndPositionItems(uint cidl, IntPtr apidl, ref POINT apt, uint dwFlags);
    }


    [ComImport,
     Guid("1af3a467-214f-4298-908e-06b03e0b39f9"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IFolderView2 : IFolderView
    {
        #region IFolderView
        new void GetCurrentViewMode([Out] out uint pViewMode);

        new int SetCurrentViewMode(uint ViewMode);
        new unsafe int GetFolder(ref Guid riid, [Out, MarshalAs(UnmanagedType.IUnknown)] out object ppv);
        new void Item(int iItemIndex, out IntPtr ppidl);
        new void ItemCount(uint uFlags, out int pcItems);
        new void Items(uint uFlags, ref Guid riid, [Out, MarshalAs(UnmanagedType.IUnknown)] out object ppv);
        new void GetSelectionMarkedItem(out int piItem);
        new void GetFocusedItem(out int piItem);
        new int GetItemPosition(IntPtr pidl, [MarshalAs(UnmanagedType.LPStruct)] out POINT ppt);
        new void GetSpacing([Out] out POINT ppt);
        new void GetDefaultSpacing(out POINT ppt);
        new void GetAutoArrange();
        new void SelectItem(int iItem, uint dwFlags);
        new void SelectAndPositionItems(uint cidl, IntPtr apidl, ref POINT apt, uint dwFlags);
        #endregion
        // IFolderView2
        void SetGroupBy(IntPtr key, bool fAscending);
        void GetGroupBy(ref IntPtr pkey, ref bool pfAscending);
        void SetViewProperty(IntPtr pidl, IntPtr propkey, object propvar);
        void GetViewProperty(IntPtr pidl, IntPtr propkey, out object ppropvar);
        void SetTileViewProperties(IntPtr pidl, [MarshalAs(UnmanagedType.LPWStr)] string pszPropList);
        void SetExtendedTileViewProperties(IntPtr pidl, in object pszPropList);
        void SetText(int iType, [MarshalAs(UnmanagedType.LPWStr)] string pwszText);
        void SetCurrentFolderFlags(uint dwMask, uint dwFlags);
        void GetCurrentFolderFlags(out uint pdwFlags);
        void GetSortColumnCount(out int pcColumns);
        void SetSortColumns(IntPtr rgSortColumns, int cColumns);
        void GetSortColumns(out IntPtr rgSortColumns, int cColumns);
        void GetItem(int iItem, ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppv);
        void GetVisibleItem(int iStart, bool fPrevious, out int piItem);
        void GetSelectedItem(int iStart, out int piItem);
        void GetSelection(bool fNoneImpliesFolder, out IShellItemArray ppsia);
        void GetSelectionState(IntPtr pidl, out uint pdwFlags);
        void InvokeVerbOnSelection([In, MarshalAs(UnmanagedType.LPWStr)] string pszVerb);

        [PreserveSig]
        int SetViewModeAndIconSize(int uViewMode, int iImageSize);

        [PreserveSig]
        int GetViewModeAndIconSize(out int puViewMode, out int piImageSize);
        void SetGroupSubsetCount(uint cVisibleRows);
        void GetGroupSubsetCount(out uint pcVisibleRows);
        void SetRedraw(bool fRedrawOn);
        void IsMoveInSameFolder();
        void DoRename();
    }
}
