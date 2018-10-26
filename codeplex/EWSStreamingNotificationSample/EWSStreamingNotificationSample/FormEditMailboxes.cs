using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace EWSStreamingNotificationSample
{
    public partial class FormEditMailboxes : Form
    {
        bool _cancel = false;
        public FormEditMailboxes()
        {
            InitializeComponent();
        }

        public string EditMailboxes(string mailboxes)
        {
            // Allow the editing of the list of mailboxes
            textBoxMailboxes.Text = mailboxes;
            _cancel = false;
            this.ShowDialog();
            if (!_cancel)
                return textBoxMailboxes.Text;
            return mailboxes;
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            _cancel = true;
            this.Hide();
        }
    }
}
