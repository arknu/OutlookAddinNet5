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

namespace OutlookAddinNet5
{
    //[ProgId("OutlookAddinNet5.TestTaskPane")]
    [Guid("EBC1A7F3-DD4B-4EBE-9DDE-4A7640E4434B")]

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
