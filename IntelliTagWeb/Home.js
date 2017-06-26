(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                // $("#template-description").text("This sample displays the selected text.");
                $('#button-text').text("Display!");
                //$('#button-desc').text("Display the selected text");

                $('#highlight-button').click(displaySelectedText);
                return;
            }

            //  $("#template-description").text("This sample highlights the longest word in the text you have selected in the document.");
            $('#GetText-button-text').text("Get Text!");
            $('#GetOOXML-button-text').text("Get OOXML!");
            $('#SetOOXML-button-text').text("Set OOXML!");
            //$('#button-desc').text("Highlights the longest word.");

            loadSampleData();

            // Add a click event handler for the highlight button.
            $('#GetText-button').click(getText);
            $('#GetOOXML-button').click(getOOXML);
            $('#SetOOXML-button').click(setOOXML);
        });
    };

    function loadSampleData() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {
            // Create a proxy object for the document body.
            var body = context.document.body;

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

    function getText() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();

            // This variable will keep the search results for the longest word.
            var searchResults;



            // Queue 
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {

                    Office.context.document.getSelectedDataAsync(
                        Office.CoercionType.Text,
                        { asyncContext: 'Some related info' },
                        function (result) {
                            if (result.status === Office.AsyncResultStatus.Failed) {
                                write('Action failed. Error: ' + result.error.message);
                            }
                            else {
                                // var output = String.format(
                                //   'Selected text: {0}<br/>Passed info: {1}',
                                // result.value, result.asyncContext);
                                var textArea = document.getElementById("dataOOXML");
                                var currentResult = result.value;
                                while (textArea.hasChildNodes()) {
                                    textArea.removeChild(textArea.lastChild);
                                    //report.innerText = "";
                                };
                                setTimeout(function () {
                                    textArea.appendChild(document.createTextNode(currentResult));
                                    //report.innerText = "The getOOXML function succeeded!";
                                }, 400);


                                //$('#result-text').text(result.value);
                            }
                        }
                    );
                    // Get the longest word from the selection.
                    /*var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');*/
                })
                .then(context.sync)
                /*.then(function () {
                    // Queue a command to highlight the search results.
                    //searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    //searchResults.items[0].font.bold = true;
                })
                .then(context.sync)*/;
        })
            .catch(errorHandler);
    }

    function getOOXML() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();

            // This variable will keep the search results for the longest word.
            var searchResults;



            // Queue 
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {

                    Office.context.document.getSelectedDataAsync(
                        Office.CoercionType.Ooxml,
                        { asyncContext: 'Some related info' },
                        function (result) {
                            if (result.status === Office.AsyncResultStatus.Failed) {
                                write('Action failed. Error: ' + result.error.message);
                            }
                            else {
                                // var output = String.format(
                                //   'Selected text: {0}<br/>Passed info: {1}',
                                // result.value, result.asyncContext);
                                var textArea = document.getElementById("dataOOXML");
                                var currentResult = result.value;
                                while (textArea.hasChildNodes()) {
                                    textArea.removeChild(textArea.lastChild);
                                    //report.innerText = "";
                                };
                                setTimeout(function () {
                                    textArea.appendChild(document.createTextNode(currentResult));
                                    //report.innerText = "The getOOXML function succeeded!";
                                }, 400);


                                //$('#result-text').text(result.value);
                            }
                        }
                    );
                    // Get the longest word from the selection.
                    /*var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');*/
                })
                .then(context.sync)
                /*.then(function () {
                    // Queue a command to highlight the search results.
                    //searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    //searchResults.items[0].font.bold = true;
                })
                .then(context.sync)*/;
        })
            .catch(errorHandler);
    }

    function setOOXML() {

        //Sets the currentOOXML variable to the current contents of the task pane text area
        var textArea = document.getElementById("dataOOXML");
        var currentOOXML = document.getElementById("dataOOXML").textContent;

        // Remove all nodes from the status Div so we have a clean space to write to
        // while (report.hasChildNodes()) {
        //   report.removeChild(report.lastChild);
        //}

        // Check whether we have OOXML in the variable
        if (currentOOXML != "") {

            // Call the setSelectedDataAsync, with parameters of:
            // 1. The Data to insert.
            // 2. The coercion type for that data.
            // 3. A callback function that lets us know if it succeeded.


            Office.context.document.setSelectedDataAsync(
                currentOOXML, { coercionType: "ooxml" },
                function (result) {
                    // Tell the user we succeeded and then clear the message after a 2 second delay
                    if (result.status == "succeeded") {
                        /*      report.innerText = "The setOOXML function succeeded!";
                              setTimeout(function () {
                                  report.innerText = "";
                              }, 2000);*/
                    }
                    else {
                        // This runs if the getSliceAsync method does not return a success flag
                        //      report.innerText = result.error.message;
                        write(result.error.message);
                        // Clear the text area just so we don't give you the impression that there's
                        // valid OOXML waiting to be inserted... 
                        while (textArea.hasChildNodes()) {
                            textArea.removeChild(textArea.lastChild);
                        }
                    }
                });
        }
        else {

            // If currentOOXML == "" then we should not even try to insert it, because
            // that is gauranteed to cause an exception, needlessly.
            //  report.innerText = "There is currently no OOXML to insert!"
            //    + " Please select some of your document and click [Get OOXML] first!";
        }
    }


    function displaySelectedText() {

        return;
        /*    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        showNotification('The selected text is:', '"' + result.value + '"');
                    } else {
                        showNotification('Error:', result.error.message);
                    }
                });*/
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
