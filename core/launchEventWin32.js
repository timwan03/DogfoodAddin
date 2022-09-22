/**
 * @file : launchEventWin32.js
 * @author : Microsoft Corporation
 */

 
 /**
  * Register all necessary LaunchEvent handlers.
  */
 Office.actions.associate("validateSendable", validateSendable);
 
 function delaySend(event) {
   var featureStatus = Office.context.roamingSettings.get("featureStatus");

    if (featureStatus & 0x00000001 /*FeatureMask for delaySend*/)
    {
        var delay = Office.context.roamingSettings.get("softBlockDelay");
        setTimeout(allowSend, delay, event);
    }
    else
    {
        allowSend(event);
    }
 }
 
function allowSend(event)
{
    event.completed({ allowEvent: true });
}

 
