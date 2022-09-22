var g_ToggleButtons = new Object(); 
const SETTINGS_SUFFIX = "-settings";
const HELP_SUFFIX = "-help";
const TOGGLE_SUFFIX = "-toggle";
// Feature List

var g_FeatureList = [
    {
        "featureId"     :   "delay-send",
        "featureLabel"  :   "Delay Send",
        "featureMask"   :   0x00000001,
    },
    {
        "featureId"     :   "customize-signature",
        "featureLabel"  :   "Customize Signature",
        "featureMask"   :   0x00000002,
    },
];

// Button Handler functions

function handleToggleClick()
{
    // When someone clicks a toggle, look at all feature that are on, and create a new mask 
    // to save to the roaming settings
    var newMask = 0;
    for (var i = 0; i < g_FeatureList.length; i++) {
        if (getToggleStatus(g_FeatureList[i].featureId))
        {
            newMask |= g_FeatureList[i].featureMask;
        }
    }

    Office.context.roamingSettings.set("featureStatus", newMask);
    Office.context.roamingSettings.saveAsync();

}

function handleHelpClick(buttonId)
{
    // $('#buttonContainer').append("Help Clicked: " + buttonId + "<br>");
}

function handleSettingsClick(buttonId)
{
    // $('#buttonContainer').append("Settings Clicked: " + buttonId + "<br>");
    
    var tabToShow = "#tab-" + buttonId;
    $("#tab-main").hide("slow");
    $(tabToShow).show();
}

function getToggleStatus(featureId)
{
    var buttonId = featureId + TOGGLE_SUFFIX;
    var toggleButton = document.querySelector("#" + buttonId);
    var toggleField = toggleButton.parentElement.querySelector(".ms-Toggle-field");
    return toggleField.classList.contains('is-selected');
}

function setToggleStatus(featureId, fToggled)
{
    if (getToggleStatus(featureId) != fToggled)
    {
        g_ToggleButtons[featureId.toString()]._toggleHandler();
    }
}

function goBackMain()
{    
    $(".tab-subpage").hide("slow");
    $("#tab-main").show("slow");
}

// Set up HTML elements functions

function setupSubpages()
{
    var backButton = '<button onClick="goBackMain();">Back</button>';
    $(".tab-subpage").append(backButton);
}

// On inital load sets the toggle switches to the correct positions 
function loadFeatureStatus()
{
    var currentStatus = Office.context.roamingSettings.get("featureStatus");
    currentStatus = currentStatus == undefined ? 0 : currentStatus;

    for (var i = 0; i < g_FeatureList.length; i++) {
        if (!!(g_FeatureList[i].featureMask & currentStatus))
        {
            setToggleStatus(g_FeatureList[i].featureId, true);
        }
    }
}

function AddFeatureButton(id, text)
{
    // Create HTML to insert feature for button
    var buttonText = '<div class="featureEntry">' + 
                        '<div id="'+ id + '-toggleParent'+'" class="ms-Toggle featureEntry-child featureEntry-toggle">' +
                            '<input type="checkbox" id="' + id + TOGGLE_SUFFIX + '" class="ms-Toggle-input" />' +
                            '<label for="'+ id +'" class="ms-Toggle-field">' +
                            '</label>' +
                        '</div>' +
                        '<div class ="featureEntry-Label featureEntry-child">' +
                            text +
                        '</div>' +
                        '<div class="rightButtons">' +
                        '<div class="featureEntry-child featureEntry-Button">' +
                            '<button class="iconButton" id="'+ id + SETTINGS_SUFFIX +'">' +
                                '<i class="ms-Icon ms-font-xl ms-Icon--Settings iconButtonIcon"></i>' + 
                            '</button>' +
                        '</div>' +
                        '<div class="featureEntry-child featureEntry-Button">' +
                            '<button class="iconButton" id="'+ id + HELP_SUFFIX +'">' +
                                '<i class="ms-Icon ms-font-xl ms-Icon--Info iconButtonIcon"></i>' + 
                            '</button>' +
                        '</div>' +
                        '</div>' +
                    '</div>';
    
    // Add button to HTML
    $('#buttonContainer').append(buttonText);

    // setup Toggle Buttons
    
    // For some reason I can't get .click to work on my toggle button, 
    // so I set up onClick handlers on the containing div 
    $("#" + id + "-toggleParent").click(function ()
    {
        handleToggleClick();
    });

    var selectorId = "#" + id + TOGGLE_SUFFIX;
    var toggleButton = document.querySelector(selectorId);
    toggleButton = toggleButton.parentElement;
    g_ToggleButtons[id.toString()] = new fabric['Toggle'](toggleButton);
    var dork = $(selectorId);

    // Setup Click Handler for Settings
    selectorId = "#" + id + SETTINGS_SUFFIX;
    $(selectorId).click(function() 
    {
        var div_id=$(this).attr("id");
        handleSettingsClick(div_id);
    });

    // Setup Click Handler for Help
    selectorId = "#" + id + HELP_SUFFIX;
    $(selectorId).click(function() 
    {
        var div_id=$(this).attr("id");
        handleHelpClick(div_id);
    });
}

// Functions for delay-send feature

function setDelay(event) {
    var delayInput = parseInt($('#delay-input').val());
    Office.context.roamingSettings.set("softBlockDelay", delayInput * 1000);
    Office.context.roamingSettings.saveAsync(function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(`Action failed with message ${result.error.message}`);
        
      } else {
        console.log(`Settings saved with status: ${result.status}`);
        $('#delay-input').attr('placeholder', delayInput);
        
      }
    });
  }

  function loadDelaySendSettingPage()
  {
    var currDelay = Office.context.roamingSettings.get("softBlockDelay");
    currDelay = currDelay == undefined ? 0 : currDelay;
    $('#delay-input').attr('placeholder', currDelay / 1000);

    $('#set-delay').click(setDelay);
  }

// Functions for set signature

  function setSignature()
  {
    Office.context.roamingSettings.set("sigTag", $('#signature-input').val());
    Office.context.roamingSettings.saveAsync();
  }

  function loadSetSignatureSettingPage()
  {
    var sigTag = Office.context.roamingSettings.get("sigTag");
    sigTag = sigTag == undefined ? "After-hours responses are not required or expected." : sigTag;

    $('#signature-input').val(sigTag);

    $('#set-signature').click(setSignature);
  }


/// ----- Office.initialize and document.ready

 $(document).ready(function(){
    $(".tab-subpage").hide();
 });


  Office.initialize = function (reason) {

    for (var i = 0; i < g_FeatureList.length; i++) {
        AddFeatureButton(g_FeatureList[i].featureId, g_FeatureList[i].featureLabel);
    }

    setupSubpages();
    loadFeatureStatus();
    loadDelaySendSettingPage();
    loadSetSignatureSettingPage();

    // $('#buttonContainer').append(JSON.stringify(getToggleStatus('delay-send')) + "<br>");
    // $('#buttonContainer').append(JSON.stringify(getToggleStatus('customize-signature')) + "<br>");


}

