function Uninstall() {

 try{
  var shell = Sys.OleObject("Shell.Application");
  var registryKey = "HKEY_LOCAL_MACHINE\\SOFTWARE\\WOW6432Node\\Thermo Scientific\\Ardia\\Ardia Platform Link";
  var valueName = "Name";
  var registryValue = Sys.OleObject("WScript.Shell").RegRead(registryKey + "\\" + valueName);
  if (registryValue === "Ardia Platform Link") {
      RunInstallationCode(shell);
      } else {
      Log.Message("Ardia already Uninstalled");
    }
  }
  catch (e)
  {
 Log.Message("Ardia already Uninstalled");
 }
}

function RunInstallationCode(shell) {
 var shell = Sys.OleObject("Shell.Application");
  var cmdPath = "cmd.exe";
  shell.ShellExecute(cmdPath, "", "", "runas", 1);
  Sys.WaitProcess("cmd.exe");
  Sys.Keys("cd /");
  Sys.Keys("[Enter]");
  Sys.Keys("cd c:\\VSTS\\3\\s\\Artifacts\\Installer");
  Sys.Keys("[Enter]");
  Sys.Keys("Ardia_Platform_Link_Setup.exe /uninstall");
  Sys.Keys("[Enter]");
  Sys.Keys("exit");
  Sys.Keys("[Enter]");
  Delay(15000)


  var uninstallConfirmationWindow = Sys.Process("Ardia_Platform_Link_Setup", 2).
  Window("#32770", "Thermo Scientific Ardia Platform Link", 1);

  if (uninstallConfirmationWindow.Exists) {
    var uninstallConfirmButton = uninstallConfirmationWindow.Window("Button", "&Yes", 1);

    if (uninstallConfirmButton.Exists) {
      uninstallConfirmButton.Click();
      Log.Message("Clicked UnInstall Button");
    } else {
      Log.Message("unInstall Button not found");
    }
  } else {
    Log.Message("Unable to find uninstallConfirmationWindow");
  }

  // Wait for the Close button to appear after uninstallation
  var closeButton = Sys.Process("Ardia_Platform_Link_Setup", 2).WPFObject("HwndSource: InstallerMainView", "Ardia Platform Link").WPFObject("InstallerMainView", "Ardia Platform Link", 1).WPFObject("Grid", "", 1).WPFObject("Grid", "", 5).WPFObject("StackPanel", "", 2).WPFObject("BtnClose");

  if (closeButton.WaitProperty("Visible", true, 1000000)) {
            try {
   var  installerWindow =Sys.Process("Ardia_Platform_Link_Setup", 2)
  .WPFObject("HwndSource: InstallerMainView", "Ardia Platform Link")
  .WPFObject("InstallerMainView", "Ardia Platform Link", 1)
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
    Log.Message("Close button clicked.");
    Log.Message("Uninstall completed.");
  } else {
    Log.Message("Close button did not appear within the time limit.");
  }

}


