//This script adds command-line support for MSI build with Visual Studio 2008. 
var msiOpenDatabaseModeTransact = 1; 

if (WScript.Arguments.Length != 1) 
{ 
    WScript.StdErr.WriteLine(WScript.ScriptName + " file"); 
    WScript.Quit(1); 
} 

WScript.Echo(WScript.Arguments(0)); 
var filespec = WScript.Arguments(0); 
var installer = WScript.CreateObject("WindowsInstaller.Installer"); 
var database = installer.OpenDatabase(filespec, msiOpenDatabaseModeTransact); 

var sql
var view
var compID = "C__E7B2925D6E19B435DF56E75F03258B57";
// update this from the 'File' table in Orca!!!

try 
{ 
    //Update InstallUISequence to support command-line parameters in interactive mode. 
    sql = "UPDATE InstallUISequence SET Condition = 'XMLPATH=\"\"' WHERE Action = 'CustomTextA_SetProperty_EDIT1'"; 
    view = database.OpenView(sql); 
    view.Execute(); 
    view.Close(); 

    //Update InstallUISequence to support command-line parameters in interactive mode. 
    sql = "UPDATE InstallUISequence SET Condition = 'FREQUENCYMINUTES=\"\"' WHERE Action = 'CustomTextA_SetProperty_EDIT2'"; 
    view = database.OpenView(sql); 
    view.Execute(); 
    view.Close(); 

    //Update InstallExecuteSequence to support command line in passive or quiet mode. 
    sql = "UPDATE InstallExecuteSequence SET Condition = 'XMLPATH=\"\"' WHERE Action = 'CustomTextA_SetProperty_EDIT1'"; 
    view = database.OpenView(sql); 
    view.Execute(); 
    view.Close(); 

    //Update InstallExecuteSequence to support command line in passive or quiet mode. 
    sql = "UPDATE InstallExecuteSequence SET Condition = 'FREQUENCYMINUTES=\"\"' WHERE Action = 'CustomTextA_SetProperty_EDIT2'"; 
    view = database.OpenView(sql); 
    view.Execute();
    view.Close();

    sql = "INSERT INTO ServiceControl (ServiceControl,Name,Event,Arguments,Wait,Component_) VALUES ('TimezoneSync','TimezoneSync',170,null,null,'C__E7B2925D6E19B435DF56E75F03258B57')";
    view = database.OpenView(sql);
    view.Execute();
    view.Close();

    database.Commit(); 
} 
catch(e) 
{ 
    WScript.StdErr.WriteLine(e); 
    WScript.Quit(1); 
} 