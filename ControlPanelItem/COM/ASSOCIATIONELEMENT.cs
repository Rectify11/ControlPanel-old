using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Rectify11.COM
{
    /// <summary>
    /// Defines information used by AssocCreateForClasses to retrieve an IQueryAssociations interface for a given file association.
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct ASSOCIATIONELEMENT
    {
        /// <summary>
        /// Where to obtain association data and the form the data is stored in. One of the following values from the ASSOCCLASS enumeration.
        /// </summary>
        public ASSOCCLASS ac;

        /// <summary>
        /// A registry key that specifies a class that contains association information.
        /// </summary>
        public IntPtr hkClass;

        /// <summary>
        /// A pointer to the name of a class that contains association information.
        /// </summary>
        [MarshalAs(UnmanagedType.LPWStr)] public string pszClass;
    }
}
