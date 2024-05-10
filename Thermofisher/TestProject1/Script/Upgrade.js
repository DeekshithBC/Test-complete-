function check(){ {

 try{
  var shell = Sys.OleObject("Shell.Application");
  var registryKey = "HKEY_LOCAL_MACHINE\\SOFTWARE\\WOW6432Node\\Thermo Scientific\\Ardia\\Ardia Platform Link";
  var valueName = "Name";
  var registryValue = Sys.OleObject("WScript.Shell").RegRead(registryKey + "\\" + valueName);
  if (registryValue === "Ardia Platform Link") {
    Upgrade();
      } else {
      Log.Message("Ardia already Uninstalled");
    }
  }
  catch (e)
  {
 Log.Message("Ardia already Uninstalled");
 }
}

function Upgrade() {
  var shell = Sys.OleObject("Shell.Application");
  var cmdPath = "cmd.exe";
  shell.ShellExecute(cmdPath, "", "", "runas", 1);
  Sys.WaitProcess("cmd.exe");
  Sys.Keys("cd /");
  Sys.Keys("[Enter]");
  Sys.Keys("cd c:\\VSTS\\3\\s\\Artifacts\\Upgrade");
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

var Upgradebutton =CloseWindow.WPFObject("BtnInstall").WPFObject("AccessText", "Upgrade", 1);
var isEnabledUpgrade = Upgradebutton.Enabled;

var Nextbutton = CloseWindow.WPFObject("BtnNext").WPFObject("AccessText", "Next", 1);
var isNextEnabled = Nextbutton.Enabled;

var Closebutton = CloseWindow.WPFObject("BtnClose").WPFObject("AccessText", "_Close", 1);
var isEnabledClose = Closebutton.Enabled;

if (isEnabledClose && isNextEnabled && !isEnabledUpgrade) {
    Log.Message("Close and Next buttons are enabled Install button is Disable.");
 
}
else {
    Log.Message(" Gettig Error:- Close and/or Next buttons are disabled. or Install button is enable");
}
  var installerWindow = Sys.Process("Ardia_Platform_Link_Setup", 2)
    .WPFObject("HwndSource: InstallerMainView", "Ardia Platform Link")
    .WPFObject("InstallerMainView", "Ardia Platform Link", 1)

  if (installerWindow.Exists) {
    var nextButton = installerWindow
      .WPFObject("Grid", "", 1)
      .WPFObject("Grid", "", 5)
      .WPFObject("StackPanel", "", 2)
      .WPFObject("BtnNext");

    if (nextButton.Exists) {
      nextButton.Click();
      Log.Message("Clicked Next Button");
    } else {
      Log.Message("NextButton Not found");
    }
  } else {
    Log.Message("installerWindow Not found");
  }

  var Upgradebutton =CloseWindow.WPFObject("BtnInstall").WPFObject("AccessText", "Upgrade", 1);
var isEnableUpgrade= Upgradebutton.Enabled;

var Nextbutton = CloseWindow.WPFObject("BtnNext").WPFObject("AccessText", "Next", 1);
var isNextEnabled = Nextbutton.Enabled;

var Closebutton = CloseWindow.WPFObject("BtnClose").WPFObject("AccessText", "_Close", 1);
var isEnabledClose = Closebutton.Enabled;
var Backbutton = CloseWindow.WPFObject("BtnBack");
var isEnableBack = Backbutton.Enabled;
  
if (isEnabledClose &&  isEnableUpgrade && isEnableBack   && !isNextEnabled) 
{
    Log.Message("Close,Back and Upgrade  are enabled Next button is Disable.");
  
}
else {
    Log.Message(" Error:-Close and/or Install and/or Back buttons are disabled. OR Next Button is Enable");
}
  
  
  if (installerWindow.Exists) {
    var UpgradeButton = installerWindow
      .WPFObject("Grid", "", 1)
      .WPFObject("Grid", "", 5)
      .WPFObject("StackPanel", "", 2)
      .WPFObject("BtnInstall")
      .WPFObject("AccessText", "Upgrade", 1);

    if (UpgradeButton.Exists) {
    UpgradeButton.Click();
      Log.Message("Clicked Upgrade Button");
    } else {
      Log.Message("Upgrade Button Not found");
    }
  } else {
    Log.Message("installerWindow Not found");
  }

  if (installerWindow.WaitProperty("Exists", true, 1100000)){
    Log.Message("Installation window appeared within the time limit.");

    var closeButtonVisible = false;
    var closeButton = installerWindow
      .WPFObject("Grid", "", 1)
      .WPFObject("Grid", "", 5)
      .WPFObject("StackPanel", "", 2)
      .WPFObject("BtnClose")
      .WPFObject("AccessText", "_Close", 1);

    closeButtonVisible = closeButton.WaitProperty("Visible", true,1100000);

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
    try {
        var iqReportButton = installerWindow
            .WPFObject("Grid", "", 1)
            .WPFObject("Grid", "", 3)
            .WPFObject("InfoErrorCtrl")
            .WPFObject("Grid", "", 1)
            .WPFObject("StackPanel", "", 1)
            .WPFObject("TextBlock", "", 10)
            .WPFObject("hyperLnkIQReport")
            .WPFObject("InlineUIContainer", "", 1)
            .WPFObject("TextBlock", "IQ Report.", 1);

        if (iqReportButton.Exists) {
            iqReportButton.Click();
            Log.Message("Clicked the 'IQ Report' button");
        } else {
            Log.Error("'IQ Report' button element not found!");
        }
    } catch (e) {
        Log.Error("Error while trying to click the 'IQ Report' button: " + e.message);
    }
      try {
        var windowCaptionPattern = "*" + "_IQ_Report.pdf - Adobe Acrobat Reader (64-bit)" + "*";
        var acrobatWindow = Sys.Process("Acrobat")
            .Window("AcrobatSDIWindow", windowCaptionPattern, 1);

        if (acrobatWindow.Exists) {
            acrobatWindow.Close();
            Log.Message("Closed the Adobe Acrobat Reader window");
        } else {
            Log.Error("Adobe Acrobat Reader window not found!");
        }
    } catch (e) {
        Log.Error("Error while trying to close the Adobe Acrobat Reader window: " + e.message);
    }
           try {
      
        // Locate the specified element using NameMapping
        var iqReportsButton = installerWindow
            .WPFObject("Grid", "", 1)
            .WPFObject("Grid", "", 3)
            .WPFObject("InfoErrorCtrl")
            .WPFObject("Grid", "", 1)
            .WPFObject("StackPanel", "", 1)
            .WPFObject("TextBlock", "", 13)
            .WPFObject("hyperLnkIQFldr")
            .WPFObject("InlineUIContainer", "", 1)
            .WPFObject("TextBlock", "Open IQ Reports folder", 1);

        // Check if the button element exists
        if (iqReportsButton.Exists) {
            // Click the button
            iqReportsButton.Click();
            Log.Message("Clicked the 'Open IQ Reports folder' button");
        } else {
            Log.Error("IQ Reports button element not found!");
        }
    } catch (e) {
        Log.Error("Error while trying to click the IQ Reports button: " + e.message);
        
    }
    var explorerWindow = Sys.Process("explorer").Window("CabinetWClass", "IQ Reports", 1);

        // Check if the window exists
        if (explorerWindow.Exists) {
            // Close the window
            explorerWindow.Close();
            Log.Message("The 'Open IQ Report Folder' button is working properly");
        } else {
            Log.Error("Window 'IQ Reports' not found!");
        }
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

      closeButton.Click();
      Log.Message("Clicked Close Button");
    } else {
      Log.Message("CloseButton did not become visible within the time limit.");
    }
  } else {
    Log.Message("Installation window did not appear within the time limit.");
  }
}}



