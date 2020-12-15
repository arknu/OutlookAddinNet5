using AddInDesignerObjects;
using Microsoft.Office.Core;
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
    public partial class  ThisAddin : _IDTExtensibility2, ICustomTaskPaneConsumer
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

            // Create the task pane using UserControl1 as the contents
            // The third parameter, when supplied, is Window object.
            myPane = myCtpFactory.CreateCTP("OutlookAddinNet5.TestTaskPane", "My Task Pane", Application.ActiveExplorer());

            //Set the dock position and show the task pane
            myPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            myPane.Visible = true;

            // CustomTaskPane.ContentControl is a reference to the control object
            myControl = (TestTaskPane)myPane.ContentControl;
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

        private ICTPFactory myCtpFactory;
        private CustomTaskPane myPane;
        private TestTaskPane myControl;
        //http://shulerent.com/2011/01/23/adding-task-panes-in-a-office-add-in-when-using-idtextensibility2/
        public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst)
        {
            // Store the CTP Factory for future use. You need this to display
            // Custom Task Panes
            myCtpFactory = CTPFactoryInst;

           

        }
    }
}
