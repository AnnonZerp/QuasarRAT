using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using xServer.Core.Packets.ServerPackets;

namespace xServer.Forms
{
    public partial class FrmCreateTask : Form
    {
        public string TaskName { get; private set; }
        public string TaskArguments { get; private set; }
        public string TaskPath { get; private set; }

        public FrmCreateTask()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            if (txtName.Text == null)
            {
                MessageBox.Show("A task name must be entered.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (txtPath.Text == null)
            {
                if (
                    MessageBox.Show(
                        "Entering an empty path will default to install location of the client. Do you wish to continue?",
                        "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                {
                    return;
                }
            }

            TaskName = txtName.Text;
            TaskArguments = txtArgs.Text;
            TaskPath = txtPath.Text;

            this.Close();
        }
    }
}
