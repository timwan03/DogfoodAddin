var g_ToggleButtons = new Object(); 

function handleHelpClick(buttonId)
{
    $('#buttonContainer').append("Help Clicked: " + buttonId + "<br>");
}

function handleSettingsClick(buttonId)
{
    $('#buttonContainer').append("Settings Clicked: " + buttonId + "<br>");
}

function getToggleStatus(buttonId)
{
    var toggleButton = document.querySelector("#" + buttonId);
    var toggleField = toggleButton.parentElement.querySelector(".ms-Toggle-field");
    return toggleField.classList.contains('is-selected');
}

function setToggleStatus(buttonId, fToggled)
{
    if (getToggleStatus(buttonId) != fToggled)
    {
        g_ToggleButtons[buttonId.toString()]._toggleHandler();
    }
}

function AddFeatureButton(id, text)
{
    // Create HTML to insert feature for button
    var buttonText = '<div class="featureEntry">' + 
                        '<div class="ms-Toggle featureEntry-child featureEntry-toggle">' +
                            '<input type="checkbox" id="' + id + '" class="ms-Toggle-input" />' +
                            '<label for="'+ id +'" class="ms-Toggle-field">' +
                            '</label>' +
                        '</div>' +
                        '<div class ="featureEntry-Label featureEntry-child">' +
                            text +
                        '</div>' +
                        '<div class="rightButtons">' +
                        '<div class="featureEntry-child featureEntry-Button">' +
                            '<button class="iconButton" id="'+ id +"settings" +'">' +
                                '<i class="ms-Icon ms-font-xl ms-Icon--Settings iconButtonIcon"></i>' + 
                            '</button>' +
                        '</div>' +
                        '<div class="featureEntry-child featureEntry-Button">' +
                            '<button class="iconButton" id="'+ id +"help" +'">' +
                                '<i class="ms-Icon ms-font-xl ms-Icon--Info iconButtonIcon"></i>' + 
                            '</button>' +
                        '</div>' +
                        '</div>' +
                    '</div>';
                            
    $('#buttonContainer').append(buttonText);

    var selectorId = "#" + id;
    var toggleButton = document.querySelector(selectorId);
    toggleButton = toggleButton.parentElement;
    g_ToggleButtons[id.toString()] = new fabric['Toggle'](toggleButton);

    // Setup Click Handler for Settings
    selectorId = "#" + id + "settings";
    var infoButton = document.querySelector(selectorId);
    $(selectorId).click(function() 
    {
        var div_id=$(this).attr("id");
        handleSettingsClick(div_id);
    });

    // Setup Click Handler for Help
    selectorId = "#" + id + "help";
    var infoButton = document.querySelector(selectorId);
    $(selectorId).click(function() 
    {
        var div_id=$(this).attr("id");
        handleHelpClick(div_id);
    });


    // $('#buttonContainer').append(JSON.stringify(toggleButton));
}

  Office.initialize = function (reason) {

    AddFeatureButton("customize-signature", "Customize Signature");
    AddFeatureButton("delay-send", "Delay Send");

    setToggleStatus('delay-send', true);
    // g_ToggleButtons['delay-send']._toggleHandler();

    $('#buttonContainer').append(JSON.stringify(getToggleStatus('delay-send')) + "<br");
    $('#buttonContainer').append(JSON.stringify(getToggleStatus('customize-signature')) + "<br");

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

