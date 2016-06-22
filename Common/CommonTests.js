function customPropertiesCallback(result)
{
  var customProps = result.value;
  customProps.set("aProp", "value");
  customProps.saveAsync(result, function(result){Office.context.mailbox.item.loadCustomPropertiesAsync($_.makeAsyncContext({testName:"setCustomProperties"}),customPropertiesSaveCallback);});
}
function customPropertiesSaveCallback(result)
{
  if(Constants.SUCCESS !== result.status)
  {
    Messages.postMessage(result.asyncContext.testName, "Failed with message: " + result.error.message);
    setResult(result.asyncContext.testName, Constants.FAILURE);
  }
  else
  {
    try
    {
      var customProps = result.value;
      Assert.areEqual(customProps.get("aProp"), "value", result.asyncContext.testName);
      customProps.remove("aProp");
      customProps.saveAsync($_.makeAsyncContext({testName:"setCustomProperties"}), successCallback);
    }
    catch(exception)
    {
      Messages.postMessage(result.asyncContext.testName, exception);
      setResult(result.asyncContext.testName, Constants.FAILURE);
    }
  }
}

displayDialogCallback = function(result)
{

  result.value.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, displayDialogCallback2);
}

displayDialogCallback2 = function(result)
{

    if (result.message == "Office.context.ui" || result.message == "window.external") {
        messagecount++;
        if (messagecount >= 2) {
            setResult("UIEXISTS", Constants.SUCCESS);
        }
    }
    else {
        Messages.postMessage("UIEXISTS", "Failed " + result.message + " was not 'Office.context.ui' or 'window.external'");
        setResult("UIEXISTS", Constants.FAILURE);
    }

  
}

var commonTestCollection =
{
  setCustomProperties: function()
  {
    Office.context.mailbox.item.loadCustomPropertiesAsync($_.makeAsyncContext({testName:"setCustomProperties"}),customPropertiesCallback);
    return false;
  },
  roamingSettings: function()
  {
    _settings = Office.context.roamingSettings;
    appRunBefore = _settings.get("featureTestApp");
    if(!appRunBefore)
    {
      Messages.postMessage("roamingSettings", "You have not successfully run this test before, please run again, if this message happens again, there is an issue in roamingSettings");
    }
    var date = Date();
    _settings.set("featureTestApp", date);

    Assert.areEqual(date, _settings.get("featureTestApp"), "roamingSettings value should be set to current date");
    _settings.saveAsync($_.makeAsyncContext({testName:"roamingSettings"}), successCallback);
    return false;
  },
  getUserIdentityTokenAsync: function()
  {
    var verify = function(result)
    {
      Assert.isTrue(result.value.length > 200, "identy token must not be empty - got " + result.value);
    }
    Office.context.mailbox.getUserIdentityTokenAsync(successCallback, {testName : "getUserIdentityTokenAsync", callbackVerification : verify});
    return false;
  },
  UIExists: function()
  {
    if (Office.context.mailbox.diagnostics.hostName == "OutlookWebApp")
    {
      Messages.postMessage("UIEXISTS", "OWA does not support the ui Object");
      setResult("UIEXISTS", Constants.SKIPPED);
      return;
    }
    try{

      messagecount = 0;
      Office.context.ui.displayDialogAsync("http://osf-agave/apps/pchan/Feature-Test/DisplayDialog/JsomChildDialog.html",
      {height:80, width:50, requireHTTPS: false, asyncContext:{testName:"UIEXISTS"}},
      displayDialogCallback);

    }catch(e)
    {
      setResult("UIEXISTS", Constants.FAILURE);
    }
  }
}
