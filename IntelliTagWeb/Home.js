(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            initialiseMessageBanner();

            // This is meant to work with Word 2016
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                showNotification("Version Not Supported Error", "This add-in needs Word 2016 to run as it makes use of the Word 1.1 API.")
                return;
            }

            LoadSettings();

//            initialiseDemoText();

            $('#GetOOXML-button-text').text("Get OOXML!");
            $('#SetOOXML-button-text').text("Set OOXML!");
            $('#SetDeontic-button-text').text("Set Deontic!");
            $('#SetTemporal-button-text').text("Set Temporal!");
            $('#SetOperational-button-text').text("Set Operational!");
            $('#RemoveAll-button-text').text("Remove all!");

            $('#GetOOXML-button').click(getOoxml_OnClick);
            $('#SetOOXML-button').click(setOoxml_OnClick);
            $('#SetDeontic-button').click(setDeontic_OnClick);
            $('#SetTemporal-button').click(setTemporal_OnClick);
            $('#SetOperational-button').click(setOperational_OnClick);
            $('#RemoveAll-button').click(removeAll_OnClick);

        });
    };

    function initialiseMessageBanner() {
        const element = document.querySelector('.ms-MessageBanner');
        messageBanner = new fabric.MessageBanner(element);
        messageBanner.hideBanner();
    }

    function initialiseDemoText() {
        // Run a batch operation against the Word object model.

        Word.run(function(context) {
                // Create a proxy object for the document body.
                const body = context.document.body;

                // Queue a commmand to clear the contents of the body.
                body.clear();
                // Queue a command to insert text into the end of the Word document body.
                body.insertText(
                    "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Nulla rutrum. Phasellus feugiat bibendum urna. Aliquam lacinia diam ac felis. In vulputate semper orci. Quisque blandit. Mauris et nibh. Aenean nulla. Mauris placerat tempor libero. \n Pellentesque bibendum.In consequat, sem molestie iaculis venenatis, orci nunc imperdiet justo, id ultricies ligula elit sit amet ante.Sed quis sem.Ut accumsan nulla vel nisi.Ut nulla enim, ullamcorper vel, semper vitae, vulputate vel, mi.Duis id magna a magna commodo interdum.",
                    Word.InsertLocation.end);

                // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
                return context.sync();
            })
            .catch(errorHandler);
    }

    function setTextArea(textValue) {
        var textArea = document.getElementById("dataOOXML");
        var currentResult = textValue;
        while (textArea.hasChildNodes()) {
            textArea.removeChild(textArea.lastChild);
        };
        setTimeout(function () {
            textArea.appendChild(document.createTextNode(currentResult));
        }, 400);
    }

    function setDeontic_OnClick() {
        setLogic(setDeonticMarker);
    }

    function setTemporal_OnClick() {
        setLogic(setTemporalMarker);
    }

    function setOperational_OnClick() {
        setLogic(setOperationalMarker);
    }

    function setLogic(setLogicFunction) {
        var functionsToExecute = [];
        functionsToExecute.push(function () { dataSelectorGetOOXML(functionsToExecute); });
        functionsToExecute.push(function () { setLogicFunction(functionsToExecute); });
        functionsToExecute.push(function () { setTextArea(dataSelectorSelectedOOXML.textBody) });

        dataSelectorGetText(functionsToExecute);
    }

//    function getText() {
//        var functionsToExecute = [];
//        functionsToExecute.push(function() { setTextArea(dataSelectorSelectedText) });
//        dataSelectorGetText(functionsToExecute);
//    }

    function getOoxml_OnClick() {
        var functionsToExecute = [];
        functionsToExecute.push(function () { setTextArea(dataSelectorSelectedOOXML.textBody) });
        dataSelectorGetOOXML(functionsToExecute);
    }

    function setOoxml_OnClick() {

        var functionsToExecute = [];
        dataSelectorSetOOXML(OOXML_SOURCE.TEXT_AREA, functionsToExecute);
    }

    function removeAll_OnClick() {
        
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
