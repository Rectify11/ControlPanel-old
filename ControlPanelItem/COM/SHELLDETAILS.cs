using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ControlPanelItem.COM
{
    /// <summary>
    /// Reports detailed information on an item in a Shell folder.
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct SHELLDETAILS
    {
        /// <summary>
        /// The alignment of the column heading and the subitem text in the column.
        /// </summary>
        public int fmt;

        /// <summary>
        /// The number of average-sized characters in the header.
        /// </summary>
        public int cxChar;

        /// <summary>
        /// An STRRET structure that includes a string with the requested information. To convert this structure to a string, use StrRetToBuf or StrRetToStr.
        /// </summary>
        public STRRET str;
    }
}
