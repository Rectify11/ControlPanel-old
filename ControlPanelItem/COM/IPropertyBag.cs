using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ControlPanelItem.COM
{
    [ComImport]
    [Guid("55272A00-42CB-11CE-8135-00AA004BB851")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropertyBag
    {
        [PreserveSig]
        int Read(
          [In, MarshalAs(UnmanagedType.LPWStr)] string pszPropName,
          [Out, MarshalAs(UnmanagedType.Struct)] out object pVar,
          [In] IErrorLog pErrorLog
        );

        [PreserveSig]
        int Write(
          [In, MarshalAs(UnmanagedType.LPWStr)] string pszPropName,
          [In, MarshalAs(UnmanagedType.Struct)] ref object pVar
        );
    }
    [System.Runtime.InteropServices.InterfaceType(1)]
    [System.Runtime.InteropServices.Guid("3127CA40-446E-11CE-8135-00AA004BB851")]
    public interface IErrorLog
    {
        void AddError(string pszPropName, ref System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo);
    }

}
