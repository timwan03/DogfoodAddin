/**
 * @file : launchEventWin32.js
 * @author : Microsoft Corporation
 */

 
 /**
  * Register all necessary LaunchEvent handlers.
  */
  if (Office.actions) {
    Office.actions.associate("delaySend", delaySend);
    Office.actions.associate("addDogfoodSignature", addDogfoodSignature);
  }

 function getFeatureStatus()
 {
    var featureStatus = Office.context.roamingSettings.get("featureStatus");
    featureStatus = featureStatus == undefined ? 0 : featureStatus;
    return featureStatus;
 }

/**
 * 
 * Delay Send Functions
 */

 function delaySend(event) {
   var featureStatus = getFeatureStatus();

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

 /**
 * 
 * Customize Signature Functions
 */

function addDogfoodSignature(eventObj)
{
  let logoContent = "iVBORw0KGgoAAAANSUhEUgAAACIAAAAiCAYAAAA6RwvCAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAIISURBVFhH7Ze9bxNBEMXfbC5OkyIUSO44KUmNS6C6dEATu0JIUQiICkUi1BQxVaAiBXQgPgoiKuIyNCQNSun8AUgXKR2WsCkiBNwOc9aAyZH7QPZe5V/hm7eW5ae9N3ezhKJsdG7A8BpANV0ZEm7Lxy4qXhP3zvSKGXnY2QThrqoRI4YqXjChKp2NL77sxJYqB1AVkf1mVKVjbF0rl9TzjcDOaOEQqhUwQnGonFMwrJ/bIDqvygkFdkSY8gIwH6hywp8d8S+s+IbsoiHMsEX4af/1K/1qwKOOBJf//znSPfLpuHtO1an0jcxdWl6XSzOuf8OMMKKfjfDjm6EzYm69bRIh/o9UzOzFpUW5njARIz/0PXgvVDrHENEDrU+jJkbLeI7EYc3uBiIzondLNsW6pgTGRpKMjSQZG0kyNpKkFCOWkPPi5FY5O/L82jaDD1X9Q0T8uLRbYyNTT5qxQA9MN/Hs+h5Vr97flamE/5qRTnA8XQ2/np0PVRaH5bzy9HJL1YDbW4GxJiBQGNnJd3jZ6MXLhNUdMZFmYzhkuGpjsrKAzYWuLqUik6EbEzEyXNXw4/sHlZk4z0jfzJ33uTNNOWElG2iVimFOb6uRQZQbdiMB2dbaCRLYHrxKbk4MvKl1aRt3hyemNemafotmMWiY1Z343NGQhZEcLeWWtyBPTDy5sqdLGQC/AJM9h+Epch8hAAAAAElFTkSuQmCC";
  let logoName = "op-logo.png";
  var tagline = Office.context.roamingSettings.get("sigTag");
  let afterHoursDisclaimer = "";
  let today = new Date();
  let time = today.getHours();
  let day = today.getDay();
  var featureStatus = getFeatureStatus();

  if (!(featureStatus & 0x00000002) /*FeatureMask for signature*/)
  {
      eventObj.completed();
      return;
  }

  if (day == 0 || day == 6 || time < 8 || time > 16) {
    afterHoursDisclaimer += "<br/>";
    afterHoursDisclaimer += "<span style='font-size:7.0pt'>"+ tagline +"</span>";
  }

  let signature = "";
  signature += "<table>";
  signature +=   "<tr>";
  signature +=     "<td style='border-right: 1px solid #888888; padding-right: 5px;'><img src='cid:" + logoName + "' alt='MS Logo' width='24' height='24' /></td>";
  signature +=     "<td style='padding-left: 5px;'>" + Office.context.mailbox.userProfile.displayName + afterHoursDisclaimer + "</td>";
  signature +=   "</tr>";
  signature += "</table>";

  Office.context.mailbox.item.addFileAttachmentFromBase64Async(logoContent, logoName, { isInline: true }, function(result) {
    Office.context.mailbox.item.body.setSignatureAsync
    (
      signature,
      {
          "coercionType": "html",
          "asyncContext" : eventObj
      },
      function (asyncResult)
      {
          asyncResult.asyncContext.completed({ "key00" : "val00" });
      }
    );
  });
}



