using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ControlPanelItem.COM
{
    //todo tidy up and name properly.
    [StructLayout(LayoutKind.Sequential)]
    internal struct ICONINFO
    {
        public bool IsIcon;
        public int xHotspot;
        public int yHotspot;
        public IntPtr MaskBitmap;
        public IntPtr ColorBitmap;
    }


}
