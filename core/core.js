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

function handleToggleClick(buttonId)
{
    $('#buttonContainer').append("Toggle Clicked: " + buttonId + "<br>");

}

function handleHelpClick(buttonId)
{
    $('#buttonContainer').append("Help Clicked: " + buttonId + "<br>");
}

function handleSettingsClick(buttonId)
{
    $('#buttonContainer').append("Settings Clicked: " + buttonId + "<br>");
    
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
    var selectorId = "#" + id + TOGGLE_SUFFIX;
    $("#" + id + "-toggleParent").click(function ()
    {
        var div_id=$(this).attr("id");
        handleToggleClick(div_id);
    });

    var toggleButton = document.querySelector(selectorId);
    toggleButton = toggleButton.parentElement;
    g_ToggleButtons[id.toString()] = new fabric['Toggle'](toggleButton);
    var dork = $(selectorId);
    // $(selectorId).click(handleToggleClick);

  

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


    // $('#buttonContainer').append(JSON.stringify(toggleButton));
}

 $(document).ready(function(){
    $(".tab-subpage").hide();
 });


  Office.initialize = function (reason) {

    for (var i = 0; i < g_FeatureList.length; i++) {
        AddFeatureButton(g_FeatureList[i].featureId, g_FeatureList[i].featureLabel);
    }

//    AddFeatureButton("customize-signature", "Customize Signature");
//    AddFeatureButton("delay-send", "Delay Send");
    // g_ToggleButtons['delay-send']._toggleHandler();


    setupSubpages();
    loadFeatureStatus();
//    setToggleStatus('delay-send', true);

    $('#buttonContainer').append(JSON.stringify(getToggleStatus('delay-send')) + "<br");
    $('#buttonContainer').append(JSON.stringify(getToggleStatus('customize-signature')) + "<br");

/*
    $('#himom').click(function()
    {
        $('#buttonContainer').append("Toggle Clicked: ");
    });
*/
    /*
    var ToggleElements = document.querySelectorAll(".ms-Toggle");
    for (var i = 0; i < ToggleElements.length; i++) {
      new fabric['Toggle'](ToggleElements[i]);
    }

    $('#set-delay').html("Lois");

    var ButtonElements = document.querySelectorAll(".ms-Button");
  for (var i = 0; i < ButtonElements.length; i++) {
    new fabric['Button'](ButtonElements[i], function() {
      // Insert Event Here
    });
    }
    */
  

}

