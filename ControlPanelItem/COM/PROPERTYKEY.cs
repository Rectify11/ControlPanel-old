using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Rectify11.COM
{
    /// <summary>
    /// Specifies the FMTID/PID identifier that programmatically identifies a property. Replaces <see cref="SharpShell.Interop.SHCOLUMNID"/>.
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct PROPERTYKEY
    {
        /// <summary>
        /// A unique GUID for the property.
        /// </summary>
        public Guid fmtid;

        /// <summary>
        /// A property identifier (PID). This parameter is not used as in SHCOLUMNID. It is recommended that you set this value to PID_FIRST_USABLE. Any value greater than or equal to 2 is acceptable.
        /// </summary>
        public uint pid;
    }
}
