var dataSelectorSelectedText;
var dataSelectorSelectedOOXML;
//var dataSelectorStructuredOOXML;

function dataSelectorGetText(functionsToExecute) {
    return dataSelectorGetData(Office.CoercionType.Text, functionsToExecute);
}

function dataSelectorGetOOXML(functionsToExecute) {
    return dataSelectorGetData(Office.CoercionType.Ooxml, functionsToExecute);
}

function structureOOXML() {
    const indexBegin = dataSelectorSelectedOOXML.indexOf("<w:body");
    const indexEnd = dataSelectorSelectedOOXML.indexOf("</w:body>") + 9;
    const documentBegin = dataSelectorSelectedOOXML.substring(0, indexBegin);
    const documentEnd = dataSelectorSelectedOOXML.substring(indexEnd, dataSelectorSelectedOOXML.length);
    const textBody = dataSelectorSelectedOOXML.substring(indexBegin, indexEnd);
    dataSelectorSelectedOOXML = { documentBegin: documentBegin, textBody: textBody, documentEnd: documentEnd };
//    return dataSelectorStructuredOOXML;
}

function dataSelectorGetData(coercionType, functionsToExecute) {
    Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();

            // Queue 
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                    .then(function () {

                        Office.context.document.getSelectedDataAsync(
                            coercionType,
                            { asyncContext: "Some related info" },
                            function (result) {
                                if (result.status === Office.AsyncResultStatus.Failed) {
                                    write("Action failed. Error: " + result.error.message);
                                }
                                else {
                                    if (coercionType === Office.CoercionType.Text) {
                                        dataSelectorSelectedText = result.value;
                                    }
                                    else
                                        if (coercionType === Office.CoercionType.Ooxml) {
                                            dataSelectorSelectedOOXML = result.value;
                                            structureOOXML();
                                        }

                                    if (functionsToExecute.length > 0) {
                                        // Remove and execute the first function on the queue
                                        (functionsToExecute.shift())();
                                    }

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

function errorHandler(error) {
    // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
    showNotification("Error:", error);
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}