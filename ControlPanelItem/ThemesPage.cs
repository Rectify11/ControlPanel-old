using Rectify11.COM;
using Rectify11.SharpShell;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Security.Cryptography;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using System.Diagnostics;

namespace Rectify11
{
    [ComVisible(true)]
    [Guid("0a852434-9b22-36d7-9985-478ccf000690")]

    public class ThemesPage : IShellFolder2, IPersistFolder2, IExplorerPaneVisibility, IShellView, COM.IServiceProvider, IFolderView
    {
        /// <summary>
        /// The absolute ID list of the folder. This is provided by IPersistFolder.
        /// </summary>
        private IdList? idListAbsolute;
        private IShellBrowser? shellBrowser;
        private ShellNamespacePage customView = new ThemesPageUI();

        #region IPersistFolder2 implementation
        public int GetClassID(out Guid pClassID)
        {
            Logger.Log("GetClassID");
            //  Set the class ID to the server id.
            pClassID = typeof(ThemesPage).GUID;
            return WinError.S_OK;
        }

        public int Initialize(IntPtr pidl)
        {
            Logger.Log("Initialize");
            //  Store the folder absolute pidl.
            idListAbsolute = PidlManager.PidlToIdlist(pidl);
            return WinError.S_OK;
        }

        public int GetCurFolder(out IntPtr ppidl)
        {
            Logger.Log("GetCurFolder");
            //  Return null if we're not initialised.
            if (idListAbsolute == null)
            {
                ppidl = IntPtr.Zero;
                return WinError.S_FALSE;
            }
            //  Return the pidl.
            ppidl = PidlManager.IdListToPidl(idListAbsolute);
            return WinError.S_OK;
        }
        #endregion
        #region IShellFolder2 implementation
        public int ParseDisplayName(IntPtr hwnd, IntPtr pbc, string pszDisplayName, ref uint pchEaten, out IntPtr ppidl, ref SFGAO pdwAttributes)
        {
            Logger.Log("ParseDisplayName does not have an implemenation");
            ppidl = IntPtr.Zero;
            return WinError.E_NOTIMPL;
        }

        public int EnumObjects(IntPtr hwnd, SHCONTF grfFlags, out IEnumIDList ppenumIDList)
        {
            //We don't have any sub-objects for now
            ppenumIDList = null;
            Logger.Log("EnumObjects");
            return WinError.S_FALSE;
        }

        public int BindToObject(IntPtr pidl, IntPtr pbc, [In] ref Guid riid, out IntPtr ppv)
        {
            Logger.Log("BindToObject does not have an implemenation");
            ppv = IntPtr.Zero;
            return WinError.E_NOTIMPL;
        }

        public int BindToStorage(IntPtr pidl, IntPtr pbc, [In] ref Guid riid, out IntPtr ppv)
        {
            Logger.Log("BindToStorage does not have an implemenation");
            ppv = IntPtr.Zero;
            return WinError.E_NOTIMPL;
        }

        public int CompareIDs(IntPtr lParam, IntPtr pidl1, IntPtr pidl2)
        {
            Logger.Log("CompareIDs does not have an implemenation");
            return WinError.E_NOTIMPL;
        }

        public int CreateViewObject(IntPtr hwndOwner, [In] ref Guid riid, out IntPtr ppv)
        {
            //  Before the contents of the folder are displayed, the shell asks for an IShellView.
            //  This function is also called to get other shell interfaces for interacting with the
            //  folder itself.
            Logger.Log("CreateViewObject");
            if (riid == typeof(IShellView).GUID)
            {
                ppv = Marshal.GetComInterfaceForObject(this, typeof(IShellView));
                return WinError.S_OK;
            }
            else
            {
                Logger.Log("CreateViewObject unknown: " + riid);
                ppv = IntPtr.Zero;
                return WinError.E_NOINTERFACE;
            }
        }

        public int GetAttributesOf(uint cidl, IntPtr apidl, ref SFGAO rgfInOut)
        {
            Logger.Log("GetAttributesOf does not have an implemenation");
            return WinError.E_NOTIMPL;
        }

        public int GetUIObjectOf(IntPtr hwndOwner, uint cidl, IntPtr apidl, [In] ref Guid riid, uint rgfReserved, out IntPtr ppv)
        {
            Logger.Log("GetUIObjectOf does not have an implemenation");
            ppv = IntPtr.Zero;
            return WinError.E_NOTIMPL;
        }

        public int GetDisplayNameOf(IntPtr pidl, SHGDNF uFlags, out STRRET pName)
        {
            Logger.Log("GetDisplayNameOf does not have an implemenation");
            pName = new STRRET();
            return WinError.E_NOTIMPL;
        }

        public int SetNameOf(IntPtr hwnd, IntPtr pidl, string pszName, SHGDNF uFlags, out IntPtr ppidlOut)
        {
            Logger.Log("SetNameOf does not have an implemenation");
            ppidlOut = IntPtr.Zero;
            return WinError.E_NOTIMPL;
        }

        public int GetDefaultSearchGUID(out Guid pguid)
        {
            Logger.Log("GetDefaultSearchGUID does not have an implemenation");
            pguid = Guid.Empty;
            return WinError.E_NOTIMPL;
        }

        public int EnumSearches(out IEnumExtraSearch ppenum)
        {
            Logger.Log("EnumSearches does not have an implemenation");
            throw new NotImplementedException();
        }

        public int GetDefaultColumn(uint dwRes, out uint pSort, out uint pDisplay)
        {
            Logger.Log("GetDefaultColumn does not have an implemenation");
            pSort = 0;
            pDisplay = 0;
            return WinError.E_NOTIMPL;
        }

        public int GetDefaultColumnState(uint iColumn, out SHCOLSTATEF pcsFlags)
        {
            Logger.Log("GetDefaultColumnState does not have an implemenation");
            pcsFlags = SHCOLSTATEF.SHCOLSTATE_SECONDARYUI;
            return WinError.E_NOTIMPL;
        }

        public int GetDetailsEx(IntPtr pidl, PROPERTYKEY pscid, out object pv)
        {
            Logger.Log("GetDetailsEx does not have an implemenation");
            pv = null;
            return WinError.E_NOTIMPL;
        }

        public int GetDetailsOf(IntPtr pidl, uint iColumn, out SHELLDETAILS psd)
        {
            Logger.Log("GetDetailsOf does not have an implemenation");
            psd = new SHELLDETAILS();
            return WinError.E_NOTIMPL;
        }

        public int MapColumnToSCID(uint iColumn, out PROPERTYKEY pscid)
        {
            Logger.Log("MapColumnToSCID does not have an implemenation");
            pscid = new();
            return WinError.E_NOTIMPL;
        }


        #endregion
        #region IShellView implemenation
        public int GetWindow(out IntPtr phwnd)
        {
            phwnd = customView.Handle;
            Logger.Log("GetWindow");
            return WinError.S_OK;
        }
        public int AddPropertySheetPages(long dwReserved, ref IntPtr lpfn, IntPtr lparam)
        {
            Logger.Log("AddPropertySheetPages does not have an implemenation");
            return WinError.S_OK;
        }
        public int ContextSensitiveHelp(bool fEnterMode)
        {
            Logger.Log("ContextSensitiveHelp does not have an implemenation");
            return WinError.S_OK;
        }

        public int TranslateAcceleratorA(MSG lpmsg)
        {
            //  TODO
            Logger.Log("TranslateAcceleratorA");
            return WinError.S_OK;
        }

        public int EnableModeless(bool fEnable)
        {
            Logger.Log("EnableModeless does not have an implemenation");
            return WinError.S_OK;
        }

        public int UIActivate(SVUIA_STATUS uState)
        {
            //  TODO
            Logger.Log("IShellView::UIActivate");
            return WinError.S_OK;
        }

        public int Refresh()
        {
            Logger.Log("Refresh does not have an implemenation");
            return WinError.S_OK;
        }

        public int CreateViewWindow([In, MarshalAs(UnmanagedType.Interface)] IShellView psvPrevious, [In] ref FOLDERSETTINGS pfs, [In, MarshalAs(UnmanagedType.Interface)] IShellBrowser psb, [In] ref RECT prcView, [In, Out] ref IntPtr phWnd)
        {
            Logger.Log("CreateViewWindow");
            //  Store the shell browser.
            shellBrowser = psb;
            customView.Browser = psb;
            //  Resize the custom view.
            customView.Bounds = new Rectangle(prcView.left, prcView.top, prcView.Width(), prcView.Height());
            customView.Visible = true;
            customView.Display();
            //  Set the handle to the handle of the custom view.
            phWnd = customView.Handle;

            //  Set the custom view to be a child of the shell browser.
            IntPtr parentWindowHandle;
            psb.GetWindow(out parentWindowHandle);
            User32.SetParent(phWnd, parentWindowHandle);

            ////Set the site
            //int hr = IUnknown_SetSite(psb, this);
            //if (hr != 0)
            //{
            //    Logger.Log("IUnknown_SetSite failed with " + new Win32Exception(hr).Message);
            //}

            //  TODO: finish this function off.
            return WinError.S_OK;
        }

        public int DestroyViewWindow()
        {
            Logger.Log("DestroyViewWindow");
            //  Hide the view window, remove it from the parent.
            customView.Visible = false;
            User32.SetParent(customView.Handle, IntPtr.Zero);

            //  Dispose of the view window.
            customView.Dispose();

            //  And we're done.
            return WinError.S_OK;
        }

        public int GetCurrentInfo(out FOLDERSETTINGS pfs)
        {
            Logger.Log("GetCurrentInfo called");
            pfs = new FOLDERSETTINGS { fFlags = 0, ViewMode = FOLDERVIEWMODE.FVM_AUTO };
            return WinError.S_OK;
        }

       

        public int SaveViewState()
        {
            //save the settings here
            Logger.Log("SaveViewState");
            return WinError.S_OK;
        }

        public int SelectItem(IntPtr pidlItem, _SVSIF uFlags)
        {
            Logger.Log("SelectItem does not have an implemenation");
            return WinError.S_OK;
        }

        public int GetItemObject(_SVGIO uItem, ref Guid riid, ref IntPtr ppv)
        {
            ppv = IntPtr.Zero;
            return WinError.E_NOTIMPL;
        }
        #endregion
        #region IFolderView implementation
        public void GetCurrentViewMode([Out] out uint pViewMode)
        {
            pViewMode = 0;
        }

        public int SetCurrentViewMode(uint ViewMode)
        {
            return WinError.E_NOTIMPL;
        }

        public unsafe int GetFolder(ref Guid riid, [Out, MarshalAs(UnmanagedType.Interface)] out object ppv)
        {
            Logger.Log("GetFolder() with " + riid);
            if (riid == typeof(IExplorerPaneVisibility).GUID)
            {
                Logger.Log("GetFolder() with IExplorerPaneVisibility");
                ppv = Marshal.GetComInterfaceForObject(this, typeof(IExplorerPaneVisibility));
                return WinError.S_OK;
            }
            else
            {
                Logger.Log("GetFolder() no such interface: " + riid.ToString());
                ppv = IntPtr.Zero;
                return WinError.E_NOINTERFACE;
            }
        }

        public void Item(int iItemIndex, out IntPtr ppidl)
        {
            Logger.Log("Item not implemented");
            throw new NotImplementedException();
        }

        public void ItemCount(uint uFlags, out int pcItems)
        {
            Logger.Log("ItemCount not implemented");
            throw new NotImplementedException();
        }

        public void Items(uint uFlags, ref Guid riid, [MarshalAs(UnmanagedType.IUnknown), Out] out object ppv)
        {
            Logger.Log("Items not implemented");
            throw new NotImplementedException();
        }

        public void GetSelectionMarkedItem(out int piItem)
        {
            Logger.Log("GetSelectionMarkedItem not implemented");
            throw new NotImplementedException();
        }

        public void GetFocusedItem(out int piItem)
        {
            Logger.Log("GetFocusedItem not implemented");
            throw new NotImplementedException();
        }

        public unsafe int GetItemPosition(IntPtr pidl, [MarshalAs(UnmanagedType.LPStruct)] out POINT ppt)
        {
            ppt = new();
            return WinError.E_NOTIMPL;
        }

        public void GetSpacing([Out] out POINT ppt)
        {
            Logger.Log("GetSpacing not implemented");
            throw new NotImplementedException();
        }

        public void GetDefaultSpacing(out POINT ppt)
        {
            Logger.Log("GetDefaultSpacing not implemented");
            throw new NotImplementedException();
        }

        public void GetAutoArrange()
        {
            Logger.Log("GetAutoArrange not implemented");
            throw new NotImplementedException();
        }

        public void SelectItem(int iItem, uint dwFlags)
        {
            Logger.Log("SelectItem not implemented");
            throw new NotImplementedException();
        }

        public void SelectAndPositionItems(uint cidl, IntPtr apidl, ref POINT apt, uint dwFlags)
        {
            Logger.Log("SelectAndPositionItems not implemented");
            throw new NotImplementedException();
        }
        #endregion
        #region IServiceProvider implementation
        public int QueryService(ref Guid guidService, ref Guid riid, out IntPtr ppvObject)
        {
            Logger.Log("QueryService() called with " + guidService);
            if (guidService == typeof(IExplorerPaneVisibility).GUID)
            {
                ppvObject = Marshal.GetComInterfaceForObject(this, typeof(IExplorerPaneVisibility));
                Logger.Log("IExplorerPaneVisibility requested!");
                return WinError.S_OK;
            }
            Logger.Log("The service: " + guidService + " does not have an implemenation");
            IntPtr nullObj = IntPtr.Zero;
            ppvObject = nullObj;
            return WinError.E_NOINTERFACE;
        }
        #endregion

        #region IExplorerPaneVisibility implementation
        public int GetPaneState(ref Guid explorerPane, out ExplorerPaneState peps)
        {
            Logger.Log("GetPaneState() called with " + explorerPane);
            peps = ExplorerPaneState.Force | ExplorerPaneState.DefaultOff;
            return WinError.S_OK;
        }
        #endregion

        public CustomQueryInterfaceResult GetInterface(ref Guid iid, out IntPtr ppv)
        {
            if (iid == typeof(IShellFolder).GUID)
            {
                ppv = Marshal.GetComInterfaceForObject((IShellFolder)this, typeof(IShellFolder), CustomQueryInterfaceMode.Ignore);
                return CustomQueryInterfaceResult.Handled;
            }
            else if (iid == typeof(IShellFolder2).GUID)
            {
                ppv = Marshal.GetComInterfaceForObject((IShellFolder2)this, typeof(IShellFolder2), CustomQueryInterfaceMode.Ignore);
                return CustomQueryInterfaceResult.Handled;
            }
            else if (iid == typeof(IPersist).GUID)
            {
                ppv = Marshal.GetComInterfaceForObject((IPersist)this, typeof(IPersist), CustomQueryInterfaceMode.Ignore);
                return CustomQueryInterfaceResult.Handled;
            }
            else if (iid == typeof(IPersistFolder).GUID)
            {
                ppv = Marshal.GetComInterfaceForObject((IPersistFolder)this, typeof(IPersistFolder), CustomQueryInterfaceMode.Ignore);
                return CustomQueryInterfaceResult.Handled;
            }
            else if (iid == typeof(IPersistFolder2).GUID)
            {
                ppv = Marshal.GetComInterfaceForObject((IPersistFolder2)this, typeof(IPersistFolder2), CustomQueryInterfaceMode.Ignore);
                return CustomQueryInterfaceResult.Handled;
            }
            else if (iid == typeof(IExplorerPaneVisibility).GUID)
            {
                ppv = Marshal.GetComInterfaceForObject((IExplorerPaneVisibility)this, typeof(IExplorerPaneVisibility), CustomQueryInterfaceMode.Ignore);
                return CustomQueryInterfaceResult.Handled;
            }
            else if (iid == typeof(IShellView).GUID)
            {
                ppv = Marshal.GetComInterfaceForObject((IShellView)this, typeof(IShellView), CustomQueryInterfaceMode.Ignore);
                return CustomQueryInterfaceResult.Handled;
            }
            else if (iid == typeof(IFolderView).GUID)
            {
                ppv = Marshal.GetComInterfaceForObject((IFolderView)this, typeof(IFolderView), CustomQueryInterfaceMode.Ignore);
                return CustomQueryInterfaceResult.Handled;
            }
            else
            {
                Logger.Log("GetInterface: unhandled: " + iid.ToString());
                ppv = IntPtr.Zero;
                return CustomQueryInterfaceResult.NotHandled;
            }
        }
    }
}
