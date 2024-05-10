﻿function Install() {
  try {
    var shell = Sys.OleObject("Shell.Application");
    var registryKey = "HKEY_LOCAL_MACHINE\\SOFTWARE\\WOW6432Node\\Thermo Scientific\\Ardia\\Ardia Platform Link";
    var valueName = "Name";

    var registryValue = Sys.OleObject("WScript.Shell").RegRead(registryKey + "\\" + valueName);

    if (registryValue === "Ardia Platform Link") {
      Log.Message("Ardia already installed");
    }
  } catch (e) {
    RunInstallationCode(shell);
  }
}

function RunInstallationCode(shell) {{
  var shell = Sys.OleObject("Shell.Application");
  var cmdPath = "cmd.exe";
  shell.ShellExecute(cmdPath, "", "", "runas", 1);
  Sys.WaitProcess("cmd.exe");
  Sys.Keys("cd ../../");
  Sys.Keys("[Enter]");
  Sys.Keys("cd c:\\VSTS\\3\\s\\Artifacts\\Installer");
  Sys.Keys("[Enter]");
  Sys.Keys("Ardia_Platform_Link_Setup.exe /install");
  Sys.Keys("[Enter]");
  Sys.Keys("exit");
  Sys.Keys("[Enter]");
  Delay(15000);
  var confirmationDialog = Sys.Process("Ardia_Platform_Link_Setup", 2)
    .Find("WndClass", "#32770", 1);

  if (confirmationDialog.Exists) {
    // Click the "Yes" button in the confirmation dialog
    var yesButton = confirmationDialog.Window("Button", "&Yes", 1);
    yesButton.Click();
    Installltion();
  } else {
    // Skip Installltion() and log a message
    Log.Message("Almanic/UDAI Confirmation dialog does not exist. Skipping Installltion().");
    Installltion(); // Proceed to Installltion() directly
  }
  
}
function Installltion(){
  var Buttons = Sys.Process("Ardia_Platform_Link_Setup", 2)
    .WPFObject("HwndSource: InstallerMainView", "Ardia Platform Link")
    .WPFObject("InstallerMainView", "Ardia Platform Link", 1)
    .WPFObject("Grid", "", 1)
    .WPFObject("Grid", "", 5)
    .WPFObject("StackPanel", "", 2);
  var Installbutton = Buttons.WPFObject("BtnInstall");
  var isEnabledInstall = Installbutton.Enabled;

  var Nextbutton = Buttons.WPFObject("BtnNext").WPFObject("AccessText", "Next", 1);
  var isNextEnabled = Nextbutton.Enabled;

  var Closebutton = Buttons.WPFObject("BtnClose").WPFObject("AccessText", "_Close", 1);
  var isEnabledClose = Closebutton.Enabled;

  if (isEnabledClose && isNextEnabled && !isEnabledInstall) {
    Log.Message("Close and Next buttons are enabled Install button is Disable.");
  } else {
     Log.Error("Error: Close and/or Next buttons are disabled, or Install button is enabled.");
  }

  var installerWindow = Sys.Process("Ardia_Platform_Link_Setup", 2)
    .WPFObject("HwndSource: InstallerMainView", "Ardia Platform Link")
    .WPFObject("InstallerMainView", "Ardia Platform Link", 1);

    var ardiaURLTextBox = installerWindow.WPFObject("Grid", "", 1)
    .WPFObject("Grid", "", 4)
    .WPFObject("tabsctrl")
    .WPFObject("StackPanel", "", 1)
    .WPFObject("BrowseFldrCtrl")
    .WPFObject("Grid", "", 1)
    .WPFObject("HeaderedItemsControl", "Options", 1)
    .WPFObject("StackPanel", "", 1)
    .WPFObject("StackPanel", "", 1)
    .WPFObject("ArdiaURL");

// Check if the object is enabled
if (ardiaURLTextBox.Enabled) {
    Log.Message("ArdiaURLTextBOx is enabled");
} else {
    Log.Error("ArdiaURLTexBox is disabled - Test Case Failed");
   
}
  try {
    // Locate the specified checkbox using NameMapping
    var checkbox = installerWindow
      .WPFObject("Grid", "", 1)
      .WPFObject("Grid", "", 4)
      .WPFObject("tabsctrl")
      .WPFObject("StackPanel", "", 1)
      .WPFObject("BrowseFldrCtrl")
      .WPFObject("Grid", "", 1)
      .WPFObject("HeaderedItemsControl", "Options", 1)
      .WPFObject("StackPanel", "", 1)
      .WPFObject("chkboxSkipServerURL_chkBox");

    // Check if the checkbox element exists
    if (checkbox.Exists) {
      // Click the checkbox
      checkbox.Click();
      Log.Message("Clicked the 'Skip Server URL' checkbox");
    } else {
      Log.Error("'Skip Server URL' checkbox element not found!");
    }
  } catch (e) {
    Log.Error("Error while trying to click the 'Skip Server URL' checkbox: " + e.message);
  }

  if (ardiaURLTextBox.Enabled) {
    Log.Error("ArdiaURL TextBox is enabled  - Test Case Failed");
} else {
    Log.Message("ArdiaURLTextbox is disabled");
   
}
  
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

  try {
    // Locate the specified button using NameMapping
    var backButton = Sys.Process("Ardia_Platform_Link_Setup", 2)
      .WPFObject("HwndSource: InstallerMainView", "Ardia Platform Link")
      .WPFObject("InstallerMainView", "Ardia Platform Link", 1)
      .WPFObject("Grid", "", 1)
      .WPFObject("Grid", "", 5)
      .WPFObject("StackPanel", "", 2)
      .WPFObject("BtnBack");

    // Check if the button element exists
    if (backButton.Exists) {
      // Click the button
      backButton.Click();
      Log.Message("Clicked the 'Back' button");
    } else {
      Log.Error("'Back' button element not found!");
    }
  } catch (e) {
    Log.Error("Error while trying to click the 'Back' button: " + e.message);
  }

  checkbox.Click();

  // URL validation
  const desiredURL = "https://hb-instrument.cmddev.thermofisher.com";
  const ardiaURL = installerWindow.WPFObject("Grid", "", 1).WPFObject("Grid", "", 4)
    .WPFObject("tabsctrl").WPFObject("StackPanel", "", 1).WPFObject("BrowseFldrCtrl")
    .WPFObject("Grid", "", 1).WPFObject("HeaderedItemsControl", "Options", 1)
    .WPFObject("StackPanel", "", 1).WPFObject("StackPanel", "", 1).WPFObject("ArdiaURL");

  if (ardiaURL) {
    const currentURL = ardiaURL.Text;

    if (currentURL !== desiredURL) {
      ardiaURL.SetText(desiredURL);
      Delay(1000); // Add a short delay after setting the text (adjust as needed).
      Log.Message(`URL "${desiredURL}" entered successfully.`);
    }
  }

  if (installerWindow.Exists) {
    var nextButton = installerWindow
      .WPFObject("Grid", "", 1)
      .WPFObject("Grid", "", 5)
      .WPFObject("StackPanel", "", 2)
      .WPFObject("BtnNext");

    if (nextButton.Exists) {
      nextButton.Click();
      Delay(10000);
      Log.Message("Clicked Next Button");
    } else {
      Log.Message("NextButton Not found");
    }
  } else {
    Log.Message("installerWindow Not found");
  }

  var browser = Sys.Browser();
  var constantPartOfURL = "https://identity.hb-instrument.cmddev.thermofisher.com/Account/Login?ReturnUrl=%2Fdevice%3FuserCode%3D";
  var page = browser.Page("*" + constantPartOfURL + "*");
  var tfLogin = page.Panel(0).Panel(0).TfLogin("applogin");
  var shadowRoot = tfLogin.ShadowRoot(0);
  var form = shadowRoot.Panel(0).Panel(0).Panel(0).Form(0);
  var usernameTextbox = form.Panel(0).TfFormField(0).Textbox("username");
  var signinButton = form.Panel(3).Panel(0).Panel(0).TfButton("signin").ShadowRoot(0).Button(0);

  if (usernameTextbox.Exists) {
    usernameTextbox.SetText("deekshith.bc1@thermofisher.com");

    if (signinButton.Exists) {
      signinButton.Click();
      Delay(45000);
      browser.Close();
    } else {
      Log.Message("Sign-in button not found.");
    }
  } else {
    Log.Message("Username textbox not found.");
  }

  if (installerWindow.Exists) {
    var installButton = installerWindow
      .WPFObject("Grid", "", 1)
      .WPFObject("Grid", "", 5)
      .WPFObject("StackPanel", "", 2)
      .WPFObject("BtnInstall");

    if (installButton.Exists) {
      var Installbutton = Buttons.WPFObject("BtnInstall");
      var isEnabledInstall = Installbutton.Enabled;

      var Nextbutton = Buttons.WPFObject("BtnNext").WPFObject("AccessText", "Next", 1);
      var isNextEnabled = Nextbutton.Enabled;

      var Closebutton = Buttons.WPFObject("BtnClose").WPFObject("AccessText", "_Close", 1);
      var isEnabledClose = Closebutton.Enabled;
      var Backbutton = Buttons.WPFObject("BtnBack");
      var isEnableBack = Backbutton.Enabled;

      if (isEnabledClose && isEnabledInstall && isEnableBack && !isNextEnabled) {
        Log.Message("Close, Back and Install buttons are enabled, Next button is Disable.");
      } else {
        Log.Error("Error: Close and/or Install and/or Back buttons are disabled, OR Next Button is Enable");
      }

      installButton.Click();
      Log.Message("Clicked INSTALL Button");

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

        closeButtonVisible = closeButton.WaitProperty("Visible", true, 1000000);

        if (closeButtonVisible) {
          try {
            var installerWindow = Sys.Process("Ardia_Platform_Link_Setup", 2)
              .WPFObject("HwndSource: InstallerMainView", "Ardia Platform Link")
              .WPFObject("InstallerMainView", "Ardia Platform Link", 1)
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
            Log.Error("Error while trying to close the Notepad window: " + e.message);
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
          closeButton.Click();
          Log.Message("Clicked Close Button");
        } else {
          Log.Message("CloseButton did not become visible within the time limit.");
        }
      } else {
        Log.Message("Installation window did not appear within the time limit.");
      }
    } else {
      Log.Message("INSTALLButton Not found");
    }
  }
}}
