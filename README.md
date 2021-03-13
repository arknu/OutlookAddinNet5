# OutlookAddinNet5

This is a proof-of-concept for making a COM-based Outlook addin using .NET 5, no VSTO required.

**This does not currently work due to bugs in WinForms.** See https://github.com/dotnet/winforms/issues/4370

## Running
After building, you need to manually register the COM host. Open an Admin command propmt in the output folder and run `regsvr32 OutlookAddinNet5.comhost.dll`

You need to add some registry keys to make Outlook load the addin, use the file registeraddin.reg to do this.

## In Outlook
When starting Outlook, you should see two buttons on the ribbon. 

Button 1 will just show a message box. 

Button 2 is intended to open a custom task pane hosting a WinForms UserControl. **This functionality is currently bropken**
