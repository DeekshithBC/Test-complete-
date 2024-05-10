﻿function OpenPowerShell() {{
    var psScriptPath = "C:\\Deekshith\\hi.ps1";
    var process = Sys.OleObject("WScript.Shell").Exec("powershell.exe -File " + psScriptPath);

    while (process.Status == 0) Delay(100);

    var exitCode = process.ExitCode;
    var output = process.StdOut.ReadAll();

    if (exitCode === 0) {
       // Log.Message("PowerShell Output: " + output);
        // Check if the output contains the   expected message
        if (output.indexOf("'Thermo Scientific Ardia Platform Link' Uninstalled Or Not found.") === -1) {
            var UnsinstallString=output;
            RunInstallationCode()
        }
         else {
            Log.Message("PowerShell Output: " + output);
            }
    } else {
        Log.Error("PowerShell Script Failed (Exit Code: " + exitCode + ")");
        Log.Error("PowerShell Error Output: " + process.StdErr.ReadAll());
    }

function RunInstallationCode() {
 var shell = Sys.OleObject("Shell.Application");
  var cmdPath = "cmd.exe";
  shell.ShellExecute(cmdPath, "", "", "runas", 1);
  Sys.WaitProcess("cmd.exe");
  Sys.Keys("cd /");
  Sys.Keys("[Enter]");
  Sys.Keys(UnsinstallString);
  Sys.Keys("[Enter]");
 Sys.Keys("exit");
  Sys.Keys("[Enter]");
  Delay(15000)
    var confirmationDialog = Sys.Process("Ardia_Platform_Link_Setup", 2)
    .Find("WndClass", "#32770", 1);

  if (confirmationDialog.Exists) {
    // Click the "Yes" button in the confirmation dialog
    var yesButton = confirmationDialog.Window("Button", "&Yes", 1);
    yesButton.Click();
    uninstalll();
  } else {
    // Skip Installltion() and log a message
    Log.Message("Almanic/UDAI Confirmation dialog does not exist. Skipping Installltion().");
    uninstalll(); // Proceed to Installltion() directly
  }
  

function uninstalll(){
  var uninstallConfirmationWindow = 
  Sys.Process("Ardia_Platform_Link_Setup", 2).Window("#32770", "Almanac/UDAI uninstall confirmation", 1)

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

}}



}}