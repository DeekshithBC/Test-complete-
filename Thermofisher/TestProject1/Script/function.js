﻿ function dee(){var browser = Sys.Browser();
  var constantPartOfURL = "https://identity.hb-instrument.cmddev.thermofisher.com/Account/Login?ReturnUrl=%2Fdevice%3FuserCode%3D";
  var page = browser.Page("*" + constantPartOfURL + "*");
  var tfLogin =page.FindElement("#applogin");
  var shadowRoot = tfLogin.ShadowRoot(0);
  var form = shadowRoot.Panel(0).Panel(0).Panel(0).Form(0);
  var usernameTextbox = Sys.Browser().Page("https://identity.ardia-release.cmddev.thermofisher.com/Account/Login?ReturnUrl=%2Fdevice%3FuserCode%3D416025380%26showContent%3DFalse").FindElement("#applogin").ShadowRoot(0).FindElement("#username")
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
  }}