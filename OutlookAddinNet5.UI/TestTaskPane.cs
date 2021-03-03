using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookAddinNet5UI
{
    //[ProgId("OutlookAddinNet5UI.TestTaskPane")]
    [Guid("2303CE06-E1EC-406C-8051-A0ABB32E8231")]

    [ComVisible(true)]
    public partial class TestTaskPane : UserControl
    {
        public TestTaskPane()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Hello from taskpane");
        }
    }
}
