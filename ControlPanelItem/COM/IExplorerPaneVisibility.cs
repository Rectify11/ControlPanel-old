using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ControlPanelItem.COM
{
    public enum ExplorerPaneState
    {
        DoNotCare = 0x00000000,
        DefaultOn = 0x00000001,
        DefaultOff = 0x00000002,
        StateMask = 0x0000ffff,
        InitialState = 0x00010000,
        Force = 0x00020000
    }
    [ComImport,
     Guid("e07010ec-bc17-44c0-97b0-46c7c95b9edc"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IExplorerPaneVisibility
    {
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        public int GetPaneState(ref Guid explorerPane, out ExplorerPaneState peps);
    };

}
