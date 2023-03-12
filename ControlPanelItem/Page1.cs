using ControlPanelItem.COM;
using ControlPanelItem.SharpShell;
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

namespace ControlPanelItem
{
    [ComVisible(true)]
    [Guid("0a852434-9b22-36d7-9985-478ccf000690")]
    public class Page1 : IPersistFolder2,
        IPersistIDList,
        IShellFolder2,
        IShellFolder, IShellView
    {
        /// <summary>
        /// The absolute ID list of the folder. This is provided by IPersistFolder.
        /// </summary>
        private IdList idListAbsolute;
        private IShellBrowser shellBrowser;
        private ShellNamespacePage customView = new Page1UI();

        #region IPersistFolder2 implementation
        public int GetClassID(out Guid pClassID)
        {
            //  Set the class ID to the server id.
            pClassID = new Guid("0a852434-9b22-36d7-9985-478ccf000690");//namespaceExtension.ServerClsid;
            return WinError.S_OK;
        }

        public int Initialize(IntPtr pidl)
        {
            //  Store the folder absolute pidl.
            idListAbsolute = PidlManager.PidlToIdlist(pidl);
            return WinError.S_OK;
        }

        public int GetCurFolder(out IntPtr ppidl)
        {
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
        #region IPersistIDList implementation
        public int SetIDList(IntPtr pidl)
        {
            return ((IPersistFolder2)this).Initialize(pidl);
        }

        public int GetIDList(out IntPtr pidl)
        {
            return ((IPersistFolder2)this).GetCurFolder(out pidl);
        }


        #endregion
        #region IShellFolder2 implementation
        public int ParseDisplayName(IntPtr hwnd, IntPtr pbc, string pszDisplayName, ref uint pchEaten, out IntPtr ppidl, ref SFGAO pdwAttributes)
        {
            MessageBox.Show("ParseDisplayName does not have an implemenation");
            throw new NotImplementedException();
        }

        public int EnumObjects(IntPtr hwnd, SHCONTF grfFlags, out IEnumIDList ppenumIDList)
        {
            //We don't have any sub-objects for now
            ppenumIDList = null;
            return WinError.S_OK;
        }

        public int BindToObject(IntPtr pidl, IntPtr pbc, [In] ref Guid riid, out IntPtr ppv)
        {
            MessageBox.Show("BindToObject does not have an implemenation");
            throw new NotImplementedException();
        }

        public int BindToStorage(IntPtr pidl, IntPtr pbc, [In] ref Guid riid, out IntPtr ppv)
        {
            MessageBox.Show("BindToStorage does not have an implemenation");
            throw new NotImplementedException();
        }

        public int CompareIDs(IntPtr lParam, IntPtr pidl1, IntPtr pidl2)
        {
            MessageBox.Show("CompareIDs does not have an implemenation");
            throw new NotImplementedException();
        }

        public int CreateViewObject(IntPtr hwndOwner, [In] ref Guid riid, out IntPtr ppv)
        {
            //  Before the contents of the folder are displayed, the shell asks for an IShellView.
            //  This function is also called to get other shell interfaces for interacting with the
            //  folder itself.

            if (riid == typeof(IShellView).GUID)
            {
                //  Now create the actual shell view.
                try
                {
                    //  Create the view, get its pointer and return success.
                    //var shellFolderView = CreateViewWindow(this);
                    ppv = Marshal.GetComInterfaceForObject(this, typeof(IShellView));
                    return WinError.S_OK;
                }
                catch (Exception exception)
                {
                    //  Log the exception, set the view to null and fail.
                  //  Diagnostics.Logging.Error("An unhandled exception occured createing the folder view.", exception);
                    ppv = IntPtr.Zero;
                    return WinError.E_FAIL;
                }
            }
            //else if (riid == typeof(Interop.IDropTarget).GUID)
            //{
            //    ppv = IntPtr.Zero;
            //    return WinError.E_NOINTERFACE;
            //}
            //else if (riid == typeof(Interop.IFolderView).GUID)
            //{
            //    ppv = Marshal.GetComInterfaceForObject(this, typeof(IFolderView));
            //    MessageBox.Show("Created the ifolderview!");
            //    return WinError.S_OK;
            //}
            //else if (riid == typeof(IContextMenu).GUID)
            //{
            //    ppv = IntPtr.Zero;
            //    return WinError.E_NOINTERFACE;
            //}
            //else if (riid == typeof(IExtractIconA).GUID)
            //{
            //    ppv = IntPtr.Zero;
            //    return WinError.E_NOINTERFACE;
            //}
            //else if (riid == typeof(IExtractIconW).GUID)
            //{
            //    ppv = IntPtr.Zero;
            //    return WinError.E_NOINTERFACE;
            //}
            //else if (riid == typeof(IQueryInfo).GUID)
            //{
            //    ppv = IntPtr.Zero;
            //    return WinError.E_NOINTERFACE;
            //}
            //else if (riid == typeof(IShellDetails).GUID)
            //{
            //    ppv = IntPtr.Zero;
            //    return WinError.E_NOINTERFACE;
            //}
            //  TODO: we have to deal with others later.
            //  IID_ICategoryProvider
            //  IID_IExplorerCommandProvider
            else
            {
               // SharpNamespaceExtension.TheLog("CreateViewObject: " + riid);
                //  We've been asked for a com inteface we cannot handle.
                ppv = IntPtr.Zero;
                //  Importantly in this case, we MUST return E_NOINTERFACE.
                return WinError.E_NOINTERFACE;
            }
        }

        public int GetAttributesOf(uint cidl, IntPtr apidl, ref SFGAO rgfInOut)
        {
            MessageBox.Show("GetAttributesOf does not have an implemenation");
            throw new NotImplementedException();
        }

        public int GetUIObjectOf(IntPtr hwndOwner, uint cidl, IntPtr apidl, [In] ref Guid riid, uint rgfReserved, out IntPtr ppv)
        {
            MessageBox.Show("GetUIObjectOf does not have an implemenation");
            throw new NotImplementedException();
        }

        public int GetDisplayNameOf(IntPtr pidl, SHGDNF uFlags, out STRRET pName)
        {
            MessageBox.Show("GetDisplayNameOf does not have an implemenation");
            throw new NotImplementedException();
        }

        public int SetNameOf(IntPtr hwnd, IntPtr pidl, string pszName, SHGDNF uFlags, out IntPtr ppidlOut)
        {
            MessageBox.Show("SetNameOf does not have an implemenation");
            throw new NotImplementedException();
        }

        public int GetDefaultSearchGUID(out Guid pguid)
        {
            MessageBox.Show("GetDefaultSearchGUID does not have an implemenation");
            throw new NotImplementedException();
        }

        public int EnumSearches(out IEnumExtraSearch ppenum)
        {
            MessageBox.Show("EnumSearches does not have an implemenation");
            throw new NotImplementedException();
        }

        public int GetDefaultColumn(uint dwRes, out uint pSort, out uint pDisplay)
        {
            MessageBox.Show("GetDefaultColumn does not have an implemenation");
            throw new NotImplementedException();
        }

        public int GetDefaultColumnState(uint iColumn, out SHCOLSTATEF pcsFlags)
        {
            MessageBox.Show("GetDefaultColumnState does not have an implemenation");
            throw new NotImplementedException();
        }

        public int GetDetailsEx(IntPtr pidl, PROPERTYKEY pscid, out object pv)
        {
            MessageBox.Show("GetDetailsEx does not have an implemenation");
            throw new NotImplementedException();
        }

        public int GetDetailsOf(IntPtr pidl, uint iColumn, out SHELLDETAILS psd)
        {
            MessageBox.Show("GetDetailsOf does not have an implemenation");
            throw new NotImplementedException();
        }

        public int MapColumnToSCID(uint iColumn, out PROPERTYKEY pscid)
        {
            MessageBox.Show("MapColumnToSCID does not have an implemenation");
            throw new NotImplementedException();
        }


        #endregion

        #region IShellView implemenation
        public int GetWindow(out IntPtr phwnd)
        {
            phwnd = customView.Handle;
            //  TODO
            return WinError.S_OK;
        }

        public int ContextSensitiveHelp(bool fEnterMode)
        {
            MessageBox.Show("ContextSensitiveHelp does not have an implemenation");
            throw new NotImplementedException();
        }

        public int TranslateAcceleratorA(MSG lpmsg)
        {
            //  TODO
            return WinError.S_OK;
        }

        public int EnableModeless(bool fEnable)
        {
            MessageBox.Show("EnableModeless does not have an implemenation");
            throw new NotImplementedException();
        }

        public int UIActivate(SVUIA_STATUS uState)
        {
            //  TODO
            return WinError.S_OK;
        }

        public int Refresh()
        {
            MessageBox.Show("Refresh does not have an implemenation");
            throw new NotImplementedException();
        }

        public int CreateViewWindow([In, MarshalAs(UnmanagedType.Interface)] IShellView psvPrevious, [In] ref FOLDERSETTINGS pfs, [In, MarshalAs(UnmanagedType.Interface)] IShellBrowser psb, [In] ref RECT prcView, [In, Out] ref IntPtr phWnd)
        {
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

            //Set the site
            int hr = IUnknown_SetSite(psb, this);
            if (hr != 0)
            {
                MessageBox.Show("setsite failed." + hr.ToString("X") + "," + new Win32Exception().Message);
            }

            //  TODO: finish this function off.
            return WinError.S_OK;
        }

        public int DestroyViewWindow()
        {
            //  Hide the view window, remove it from the parent.
            customView.Visible = false;
            User32.SetParent(customView.Handle, IntPtr.Zero);

            //  Dispose of the view window.
            customView.Dispose();

            //  And we're done.
            return WinError.S_OK;
        }

        public int GetCurrentInfo(ref FOLDERSETTINGS pfs)
        {
            pfs = new FOLDERSETTINGS { fFlags = 0, ViewMode = FOLDERVIEWMODE.FVM_AUTO };
            return WinError.S_OK;
        }

        public int AddPropertySheetPages(long dwReserved, ref IntPtr lpfn, IntPtr lparam)
        {
            MessageBox.Show("AddPropertySheetPages does not have an implemenation");
            throw new NotImplementedException();
        }

        public int SaveViewState()
        {
            //save the settings here

            return WinError.S_OK;
        }

        public int SelectItem(IntPtr pidlItem, _SVSIF uFlags)
        {
            MessageBox.Show("SelectItem does not have an implemenation");
            throw new NotImplementedException();
        }

        public int GetItemObject(_SVGIO uItem, ref Guid riid, ref IntPtr ppv)
        {
            // Returning S_OK will cause Explorer to crash when navigating away from namespace View
            return WinError.E_NOTIMPL;
        }
        #endregion
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int IUnknown_SetSite(
         [In, MarshalAs(UnmanagedType.IUnknown)] object punk,
         [In, MarshalAs(UnmanagedType.IUnknown)] object punkSite);
    }
}
