using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ControlPanelItem.COM
{
    /// <summary>
    /// Defines a mechanism for retrieving a service object; that is, an object that provides custom support to other objects.
    /// </summary>
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("6d5140c1-7436-11ce-8034-00aa006009fa")]
    public unsafe interface IServiceProvider
    {
        /// <summary>
        /// Defines a mechanism for retrieving a service object; that is, an object that provides custom support to other objects.
        /// </summary>
        /// <param name="guidService">The unique identifier of the service (an SID).</param>
        /// <param name="riid">The unique identifier of the interface that the caller wants to receive for the service.</param>
        /// <param name="ppvObject">The address of the caller-allocated variable to receive the interface pointer of the service on successful return from this function. The caller becomes responsible for calling Release through this interface pointer when the service is no longer required.</param>
        /// <returns>Returns one of the following values.</returns>
        [PreserveSig]
        int QueryService(ref Guid guidService, ref Guid riid, out IntPtr ppv);

    };
}
