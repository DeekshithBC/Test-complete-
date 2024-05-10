
function Repair() {{

 try{
  var shell = Sys.OleObject("Shell.Application");
  var registryKey = "HKEY_LOCAL_MACHINE\\SOFTWARE\\WOW6432Node\\Thermo Scientific\\Ardia\\Ardia Platform Link";
  var valueName = "Name";
  var registryValue = Sys.OleObject("WScript.Shell").RegRead(registryKey + "\\" + valueName);
  if (registryValue === "Ardia Platform Link") {
     RunRepair();
      } else {
      Log.Message("Ardia already Uninstalled");
    }
  }
  catch (e)
  {
 Log.Message("Ardia already Uninstalled");
 }
}
function  RunRepair() {
  var shell = Sys.OleObject("Shell.Application");
  var cmdPath = "cmd.exe";
  shell.ShellExecute(cmdPath, "", "", "runas", 1);
  Sys.WaitProcess("cmd.exe");
  Sys.Keys("cd /");
  Sys.Keys("[Enter]");
  Sys.Keys("cd c:\\VSTS\\3\\s\\Artifacts\\Installer");
  Sys.Keys("[Enter]");
  Sys.Keys("Ardia_Platform_Link_Setup.exe /install");
  Sys.Keys("[Enter]");
  Sys.Keys("exit");
  Sys.Keys("[Enter]");
  Delay(15000);
  
  var CloseWindow = Sys.Process("Ardia_Platform_Link_Setup", 2)
    .WPFObject("HwndSource: InstallerMainView", "Ardia Platform Link")
    .WPFObject("InstallerMainView", "Ardia Platform Link", 1)
    .WPFObject("Grid", "", 1)
    .WPFObject("Grid", "", 5)
    .WPFObject("StackPanel", "", 2);

var Repairbutton = CloseWindow.WPFObject("btnRepair");
var isEnabledRepair = Repairbutton.Enabled;

var Uninstallbutton =CloseWindow.WPFObject("BtnUnInstall");
var isEnabledUninstall = Uninstallbutton.Enabled;

var Closebutton = CloseWindow.WPFObject("BtnClose").WPFObject("AccessText", "_Close", 1);
var isEnabledClose = Closebutton.Enabled;

if (isEnabledClose && isEnabledRepair && isEnabledUninstall ) {
    Log.Message("Close,repair,uninstall  buttons are enabled .");
 
}
else {
    Log.Message(" Gettig Error:- Close and/or Next buttons are disabled. or Install button is enable");
}
  
  var installerWindow = Sys.Process("Ardia_Platform_Link_Setup", 2)
  .WPFObject("HwndSource: InstallerMainView", "Ardia Platform Link")
  .WPFObject("InstallerMainView", "Ardia Platform Link", 1)
  
  
  if (installerWindow.Exists) {
    var RepairButton = installerWindow
     .WPFObject("Grid", "", 1)
  .WPFObject("Grid", "", 5).WPFObject("StackPanel", "", 2).WPFObject("btnRepair")
    if ( RepairButton .Exists) {
      RepairButton .Click();
        // Wait for the installation window to appear
      if (installerWindow.WaitProperty("Exists", true, 1000000)) {
        Log.Message("Installation window appeared within the time limit.");

        // Wait for the "Close" button to become visible
        var closeButtonVisible = false;
        var closeButton = installerWindow
          .WPFObject("Grid", "", 1)
          .WPFObject("Grid", "", 5)
          .WPFObject("StackPanel", "", 2)
          .WPFObject("BtnClose")
          .WPFObject("AccessText", "_Close", 1);

        closeButtonVisible = closeButton.WaitProperty("Visible", true, 900000);

        if (closeButtonVisible) { 
    
        try {
         
        var logFileButton = installerWindow
            .WPFObject("Grid", "", 1)
            .WPFObject("Grid", "", 3)
            .WPFObject("InfoErrorCtrl")
            .WPFObject("Grid", "", 1)
            .WPFObject("StackPanel", "", 1)
            .WPFObject("TextBlock", "", 11)
            .WPFObject("hyperLnkLogFile")
            .WPFObject("InlineUIContainer", "", 1)
            .WPFObject("TextBlock", "log", 1);

        if (logFileButton.Exists) {
            logFileButton.Click();
            Log.Message("Clicked the 'log' button");
        } else {
            Log.Error("'log' button element not found!");
        }
    } catch (e) {
        Log.Error("Error while trying to click the 'log' button: " + e.message);
    }
      try {
        // Close the Notepad window
        var constantPart = "Thermo_Scientific_Ardia_Platform_Link_";
        var variablePart = "*";
        var windowCaptionPattern = constantPart + variablePart + ".log - Notepad";

        var notepadWindow = Sys.Process("notepad")
            .Window("Notepad", windowCaptionPattern, 1);

        if (notepadWindow.Exists) {
            notepadWindow.Close();
            Log.Message("Closed the Notepad window");
        } else {
            Log.Error("Notepad window not found!");
        }
    } catch (e) {
        Log.Error("Error while trying to close the Notepad window: ");
    }

  
      closeButton.Click();
      Log.Message("Clicked Close Button");
    } else {
      Log.Message("CloseButton Not found");
    }
  }
}}

}}