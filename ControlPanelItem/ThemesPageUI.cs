using Rectify11.COM;
using Rectify11.SharpShell;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Rectify11
{
    public partial class ThemesPageUI : ShellNamespacePage
    {
        public ThemesPageUI()
        {
            InitializeComponent();
            comboBox1.SelectedIndex = 0;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            IntPtr pidl;
            Shell32.SHGetKnownFolderIDList(KnownFolders.FOLDERID_ControlPanelFolder, KNOWN_FOLDER_FLAG.KF_NO_FLAGS, IntPtr.Zero,
                out pidl);
            Browser.BrowseObject(pidl, COM.SBSP.SBSP_SAMEBROWSER);
        }
    }
}
