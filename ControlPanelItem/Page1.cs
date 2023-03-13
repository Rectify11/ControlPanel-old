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
using System.Security.Cryptography;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace ControlPanelItem
{
    [ComVisible(true)]
    [Guid("0a852434-9b22-36d7-9985-478ccf000690")]

    public class Page1 : IPersistFolder2,
        IPersistIDList,
        IShellFolder2,
        IShellFolder, 
        IShellView,
        //IShellView2,
        //We need to implement these 3 interfaces so we can hide the ribbon, command bar, tree view, etc
        COM.IServiceProvider,
        IExplorerPaneVisibility,
        IFolderView, IFolderView2, ICustomQueryInterface, IPropertyBag
    {
        /// <summary>
        /// The absolute ID list of the folder. This is provided by IPersistFolder.
        /// </summary>
        private IdList idListAbsolute;
        private IShellBrowser shellBrowser;
        private ShellNamespacePage customView = new Page1UI();

        public Page1()
        {
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            MessageBox.Show("Unhandled exception in rectify11 control panel: " + e.ExceptionObject.ToString());
        }

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
            shellBrowser.GetWindow(out IntPtr window);
            User32.SendMessage(window, 0x001A, 0, Marshal.StringToHGlobalUni("ShellState"));
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
                Logger.Log("IUnknown_SetSite failed with " + new Win32Exception(hr).Message);
                MessageBox.Show("setsite failed: " + new Win32Exception(hr).Message);
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
            Logger.Log("GetCurrentInfo called");
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
        #region IShellView2 implementation
        public int CreateViewWindow2(IntPtr ptr)
        {
            

            Logger.Log("CreateViewWindow2 called!");
            var lpParams = Marshal.PtrToStructure<SV2CVW2_PARAMS>(ptr);
            Logger.Log("CreateViewWindow2 called2!");
            //  Store the shell browser.
            shellBrowser = lpParams.psbOwner;
            customView.Browser = lpParams.psbOwner;
            //  Resize the custom view.
            customView.Bounds = new Rectangle(lpParams.prcView.left, lpParams.prcView.top, lpParams.prcView.Width(), lpParams.prcView.Height());
            customView.Visible = true;
            customView.Display();

            lpParams.psbOwner.SetStatusTextSB("Loading...");

            //  Set the handle to the handle of the custom view.
            //   lpParams.hwndView = customView.Handle;

            //  Set the custom view to be a child of the shell browser.
            IntPtr parentWindowHandle;
            lpParams.psbOwner.GetWindow(out parentWindowHandle);
            User32.SetParent(lpParams.hwndView, customView.Handle);

            //Set the site
            int hr = IUnknown_SetSite(lpParams.psbOwner, this);
            if (hr != 0)
            {
                Logger.Log("IUnknown_SetSite failed with " + new Win32Exception(hr).Message);
                MessageBox.Show("setsite failed: " + new Win32Exception(hr).Message);
            }
            Logger.Log("CreateViewWindow2 end");
            //  TODO: finish this function off.
            return WinError.S_OK;
        }

        public int GetView(in Guid pvid, in ulong uView)
        {
            MessageBox.Show("IShellView2::GetView not implemented");
            throw new NotImplementedException();
        }

        public int HandleRename(IntPtr pidlNew)
        {
            MessageBox.Show("IShellView2::HandleRename not implemented");
            throw new NotImplementedException();
        }

        public void SelectAndPositionItem(IntPtr pidlItem, SVSIF flags, in POINT point)
        {
            MessageBox.Show("IShellView2::SelectAndPositionItem not implemented");
            throw new NotImplementedException();
        }
        #endregion
        #region IFolderView implementation
        uint viewmode = 0;
        public void GetCurrentViewMode([Out] out uint pViewMode)
        {
            pViewMode = viewmode;
        }

        public int SetCurrentViewMode(uint ViewMode)
        {
            viewmode = ViewMode;
            return WinError.S_OK;
        }

        public int GetFolder(ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppv)
        {
            try
            {
                Logger.Log("GetFolder() with " + riid);
                if (riid == typeof(IExplorerPaneVisibility).GUID)
                {
                    Logger.Log("GetFolder() with IExplorerPaneVisibility");
                    ppv = Marshal.GetComInterfaceForObject(this, typeof(IExplorerPaneVisibility));
                    return WinError.S_OK;
                }
                else if (riid == typeof(IShellFolder).GUID)
                {
                    ppv = Marshal.GetComInterfaceForObject(this, typeof(IShellFolder));
                    return WinError.S_OK;
                }
                else
                {
                    MessageBox.Show("GetFolder() no such interface: " + riid.ToString());
                    Logger.Log("GetFolder() no such interface: " + riid.ToString());
                    ppv = IntPtr.Zero;

                    return WinError.E_NOINTERFACE;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("A fatal error has occured: " + ex.ToString());
                ppv = IntPtr.Zero;
                return WinError.E_NOINTERFACE;
            }
        }

        public void Item(int iItemIndex, out IntPtr ppidl)
        {
            MessageBox.Show("Item not implemented");
            throw new NotImplementedException();
        }

        public void ItemCount(uint uFlags, out int pcItems)
        {
            MessageBox.Show("ItemCount not implemented");
            throw new NotImplementedException();
        }

        public void Items(uint uFlags, ref Guid riid, [MarshalAs(UnmanagedType.IUnknown), Out] out object ppv)
        {
            MessageBox.Show("Items not implemented");
            throw new NotImplementedException();
        }

        public void GetSelectionMarkedItem(out int piItem)
        {
            MessageBox.Show("GetSelectionMarkedItem not implemented");
            throw new NotImplementedException();
        }

        public void GetFocusedItem(out int piItem)
        {
            MessageBox.Show("GetFocusedItem not implemented");
            throw new NotImplementedException();
        }

        public unsafe int GetItemPosition(IntPtr pidl, [MarshalAs(UnmanagedType.LPStruct)] out POINT ppt)
        {
            MessageBox.Show("GetItemPosition: executing");
            var p = new POINT() { X = 0, Y = 0 };
            ppt = p;
            MessageBox.Show("GetItemPosition: ok");
            return WinError.S_OK;
        }

        public void GetSpacing([Out] out POINT ppt)
        {
            MessageBox.Show("GetSpacing not implemented");
            throw new NotImplementedException();
        }

        public void GetDefaultSpacing(out POINT ppt)
        {
            MessageBox.Show("GetDefaultSpacing not implemented");
            throw new NotImplementedException();
        }

        public void GetAutoArrange()
        {
            MessageBox.Show("GetAutoArrange not implemented");
            throw new NotImplementedException();
        }

        public void SelectItem(int iItem, uint dwFlags)
        {
            MessageBox.Show("SelectItem not implemented");
            throw new NotImplementedException();
        }

        public void SelectAndPositionItems(uint cidl, IntPtr apidl, ref POINT apt, uint dwFlags)
        {
            MessageBox.Show("SelectAndPositionItems not implemented");
            throw new NotImplementedException();
        }
        #endregion
        #region IFolderView2 Implemenation

        public void SetGroupBy(IntPtr key, bool fAscending)
        {
            MessageBox.Show("SetGroupBy not implemented");
            throw new NotImplementedException();
        }

        public void GetGroupBy(ref IntPtr pkey, ref bool pfAscending)
        {
            MessageBox.Show("GetGroupBy not implemented");
            throw new NotImplementedException();
        }

        public void SetViewProperty(IntPtr pidl, IntPtr propkey, object propvar)
        {
            MessageBox.Show("SetViewProperty not implemented");
            throw new NotImplementedException();
        }

        public void GetViewProperty(IntPtr pidl, IntPtr propkey, out object ppropvar)
        {
            MessageBox.Show("GetViewProperty not implemented");
            throw new NotImplementedException();
        }

        public void SetTileViewProperties(IntPtr pidl, [MarshalAs(UnmanagedType.LPWStr)] string pszPropList)
        {
            MessageBox.Show("SetTileViewProperties not implemented");
            //throw new NotImplementedException();
        }

        public void SetExtendedTileViewProperties(IntPtr pidl, in IntPtr pszPropList)
        {
            //we can't pass pszPropList as a string or else crash!
            try
            {
                MessageBox.Show("SetExtendedTileViewProperties not implemented");
               // var pszPropList2 = Marshal.PtrToStringUni(pszPropList);

             //   Logger.Log("SetExtendedTileViewProperties called with " + pszPropList2);
               // MessageBox.Show("SetExtendedTileViewProperties not implemented end");
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void SetText(int iType, [MarshalAs(UnmanagedType.LPWStr)] string pwszText)
        {
            MessageBox.Show("SetText not implemented");
            throw new NotImplementedException();
        }

        public void SetCurrentFolderFlags(uint dwMask, uint dwFlags)
        {
            MessageBox.Show("SetCurrentFolderFlags not implemented");
            throw new NotImplementedException();
        }

        public void GetCurrentFolderFlags(out uint pdwFlags)
        {
            MessageBox.Show("GetCurrentFolderFlags not implemented");
            throw new NotImplementedException();
        }

        public void GetSortColumnCount(out int pcColumns)
        {
            MessageBox.Show("GetSortColumnCount not implemented");
            throw new NotImplementedException();
        }

        public void SetSortColumns(IntPtr rgSortColumns, int cColumns)
        {
            MessageBox.Show("SetSortColumns not implemented");
            throw new NotImplementedException();
        }

        public void GetSortColumns(out IntPtr rgSortColumns, int cColumns)
        {
            MessageBox.Show("GetSortColumns not implemented");
            throw new NotImplementedException();
        }

        public void GetItem(int iItem, ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppv)
        {
            MessageBox.Show("GetItem not implemented");
            throw new NotImplementedException();
        }

        public void GetVisibleItem(int iStart, bool fPrevious, out int piItem)
        {
            MessageBox.Show("GetVisibleItem not implemented");
            throw new NotImplementedException();
        }

        public void GetSelectedItem(int iStart, out int piItem)
        {
            MessageBox.Show("GetSelectedItem not implemented");
            throw new NotImplementedException();
        }

        void IFolderView2.GetSelection(bool fNoneImpliesFolder, out IShellItemArray ppsia)
        {
            MessageBox.Show("GetSelection not implemented");
            throw new NotImplementedException();
        }

        public void GetSelectionState(IntPtr pidl, out uint pdwFlags)
        {
            MessageBox.Show("GetSelectionState not implemented");
            throw new NotImplementedException();
        }

        public void InvokeVerbOnSelection([In, MarshalAs(UnmanagedType.LPWStr)] string pszVerb)
        {
            MessageBox.Show("InvokeVerbOnSelection not implemented");
            throw new NotImplementedException();
        }

        public int SetViewModeAndIconSize(int uViewMode, int iImageSize)
        {
            MessageBox.Show("SetViewModeAndIconSize not implemented");
            throw new NotImplementedException();
        }

        public int GetViewModeAndIconSize(out int puViewMode, out int piImageSize)
        {
            MessageBox.Show("GetViewModeAndIconSize not implemented");
            throw new NotImplementedException();
        }

        public void SetGroupSubsetCount(uint cVisibleRows)
        {
            MessageBox.Show("SetGroupSubsetCount not implemented");
            throw new NotImplementedException();
        }

        public void GetGroupSubsetCount(out uint pcVisibleRows)
        {
            MessageBox.Show("GetGroupSubsetCount not implemented");
            throw new NotImplementedException();
        }

        public void SetRedraw(bool fRedrawOn)
        {
            MessageBox.Show("SetRedraw not implemented");
            throw new NotImplementedException();
        }

        public void IsMoveInSameFolder()
        {
            MessageBox.Show("IsMoveInSameFolder not implemented");
            throw new NotImplementedException();
        }

        public void DoRename()
        {
            MessageBox.Show("DoRename not implemented");
            throw new NotImplementedException();
        }
        #endregion
        #region IServiceProvider implementation
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int IUnknown_SetSite(
         [In, MarshalAs(UnmanagedType.IUnknown)] object punk,
         [In, MarshalAs(UnmanagedType.IUnknown)] object punkSite);

        public int QueryService(ref Guid guidService, ref Guid riid, out IntPtr ppv)
        {
            if (guidService == typeof(IExplorerPaneVisibility).GUID)
            {
                ppv = Marshal.GetComInterfaceForObject(this, typeof(IExplorerPaneVisibility));
                MessageBox.Show("IExplorerPaneVisibility requested!");
                return WinError.S_OK;
            }
            else if (guidService == typeof(IFolderView).GUID)
            {
                ppv = Marshal.GetComInterfaceForObject(this, typeof(IFolderView));
                return WinError.S_OK;
            }
            MessageBox.Show("The service: " + guidService + " does not have an implemenation");
            ppv = IntPtr.Zero;
            return WinError.E_NOINTERFACE;
        }
        #endregion

        #region IExplorerPaneVisibility implementation
        public int GetPaneState(ref Guid explorerPane, out ExplorerPaneState peps)
        {
            Logger.Log("GetPaneState() called with " + explorerPane);
            //MessageBox.Show("GetPaneState called!");
            peps = ExplorerPaneState.Force | ExplorerPaneState.DefaultOff;
            return WinError.S_OK;
        }

        public CustomQueryInterfaceResult GetInterface(ref Guid iid, out IntPtr ppv)
        {
            Logger.Log("QueryInterface: " + iid.ToString());
            ppv = IntPtr.Zero;
            return CustomQueryInterfaceResult.NotHandled;
        }

        public int Read([In, MarshalAs(UnmanagedType.LPWStr)] string pszPropName, [MarshalAs(UnmanagedType.Struct), Out] out object pVar, [In] IErrorLog pErrorLog)
        {
            Logger.Log("Property bag read: " + pszPropName);
            throw new NotImplementedException();
        }

        public int Write([In, MarshalAs(UnmanagedType.LPWStr)] string pszPropName, [In, MarshalAs(UnmanagedType.Struct)] ref object pVar)
        {
            Logger.Log("Property bag write: " + pszPropName);
            throw new NotImplementedException();
        }



        #endregion

    }
}
