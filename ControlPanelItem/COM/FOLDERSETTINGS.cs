using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ControlPanelItem.COM
{
    /// <summary>
    /// Contains folder view information.
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct FOLDERSETTINGS
    {
        /// <summary>
        /// Folder view mode. One of the FOLDERVIEWMODE values.
        /// </summary>
        public FOLDERVIEWMODE ViewMode;

        /// <summary>
        /// A set of flags that indicate the options for the folder. This can be zero or a combination of the FOLDERFLAGS values.
        /// </summary>
        public FOLDERFLAGS fFlags;
    }
}
