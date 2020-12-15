using AddInDesignerObjects;
using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddinNet5
{
    [ComVisible(true)]
    [Guid("011150D8-70D0-42C2-B00A-3A1290399760")]
    [ProgId("OutlookAddinNet5")]
    public partial class  ThisAddin : _IDTExtensibility2
    {
        public Outlook.Application Application { get; set; }

        /// <summary>
        /// A simple command
        /// </summary>
        public void Command1()
        {
            MessageBox.Show(
                "Hello from command 1!",
                "VisioComAddinNet5");
        }

        /// <summary>
        /// A command to demonstrate conditionally enabling/disabling.
        /// The command gets enabled only when a shape is selected
        /// </summary>
        public void Command2()
        {
           
            MessageBox.Show(
                string.Format("Hello from (conditional) command 2!"));
        }

        /// <summary>
        /// Callback called by the UI manager when user clicks a button
        /// Should do something meaningful when corresponding action is called.
        /// </summary>
        public void OnCommand(string commandId)
        {
            switch (commandId)
            {
                case "Command1":
                    Command1();
                    return;

                case "Command2":
                    Command2();
                    return;
            }
        }

        /// <summary>
        /// Callback called by UI manager.
        /// Should return if corresponding command should be enabled in the user interface.
        /// By default, all commands are enabled.
        /// </summary>
        public bool IsCommandEnabled(string commandId)
        {
            switch (commandId)
            {
                case "Command1":    // make command1 always enabled
                    return true;

                case "Command2":    // make command2 enabled only if a drawing is opened
                    return true;
                default:
                    return true;
            }
        }

        /// <summary>
        /// Callback called by UI manager.
        /// Should return if corresponding command (button) is pressed or not (makes sense for toggle buttons)
        /// </summary>
        public bool IsCommandChecked(string command)
        {
            return false;
        }
        /// <summary>
        /// Callback called by UI manager.
        /// Returns a label associated with given command.
        /// We assume for simplicity taht command labels are named simply named as [commandId]_Label (see resources)
        /// </summary>
        public string GetCommandLabel(string command)
        {
            return Properties.Resources.ResourceManager.GetString(command + "_Label");
        }

        /// <summary>
        /// Returns a bitmap associated with given command.
        /// We assume for simplicity that bitmap ids are named after command id.
        /// </summary>
        public Bitmap GetCommandBitmap(string id)
        {
            var obj = Properties.Resources.ResourceManager.GetObject(id);
            return obj as Bitmap;
        }

        internal void UpdateUI()
        {
            UpdateRibbon();
        }



        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            this.Application = (Outlook.Application)Application;
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            
        }
        
        public void OnStartupComplete(ref Array custom)
        {
            
        }

        public void OnBeginShutdown(ref Array custom)
        {
            
        }
    }
}
