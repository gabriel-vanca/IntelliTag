var dataSelectorSelectedText;
var dataSelectorSelectedOOXML;

const OOXML_SOURCE = {
    TEXT_AREA: 'text_area',
    MARKER_EDITOR: 'marker_editor'
}

function dataSelectorGetText(functionsToExecute) {
    return dataSelectorGetData(Office.CoercionType.Text, functionsToExecute);
}

function dataSelectorGetOOXML(functionsToExecute) {
    return dataSelectorGetData(Office.CoercionType.Ooxml, functionsToExecute);
}

function dataSelectorSetOOXML(dataSource, functionsToExecute) {
    return dataSelectorSetData(Office.CoercionType.Ooxml, dataSource, functionsToExecute);
}

// Gets the OOXML contents of the Word document body and
// puts the OOXML into a textarea in the add-in.
function dataSelectorGetData(coercionType, functionsToExecute) {

    // Run a batch operation against the Word Javascript object model.
    Word.run(function(context) {

            // Create a proxy object for the document body.
            var range = context.document.getSelection();

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function() {

                    // Queue 
                    context.load(range, 'text');

                    var readData;
                    if (coercionType === Office.CoercionType.Ooxml) {
                        // Queue a commmand to get the OOXML contents of the body.
                        readData = range.getOoxml();
                    } else if (coercionType === Office.CoercionType.Text) {
                        readData = range;
                        // Queue 
                        context.load(readData, 'text');
                    }
                
                    // Synchronize the document state by executing the queued commands, 
                    // and return a promise to indicate task completion.
                    return context.sync().then(function() {

                        // Update the status message.
                        //  setTimeout(function () {
                        if (coercionType === Office.CoercionType.Text) {
                            dataSelectorSelectedText = readData.text;
                        } else if (coercionType === Office.CoercionType.Ooxml) {
                            dataSelectorSelectedOOXML = readData.value;
                            structureOOXML();
                        }

                        if (functionsToExecute.length > 0) {
                            // Remove and execute the first function on the queue
                            (functionsToExecute.shift())();
                        }
                        //  }, 400);

                    });
                })
                .catch(errorHandler);
        })
        .catch(errorHandler);
}

//function _dataSelectorGetData(coercionType, functionsToExecute) {
//    Word.run(function (context) {
//            // Queue a command to get the current selection and then
//            // create a proxy range object with the results.
//            var range = context.document.getSelection();
//
//            // Queue 
//            context.load(range, 'text');
//
//            // Synchronize the document state by executing the queued commands
//            // and return a promise to indicate task completion.
//            return context.sync()
//                    .then(function () {
//
//                        Office.context.document.getSelectedDataAsync(
//                            coercionType,
//                            { asyncContext: "Some related info" },
//                            function (result) {
//                                if (result.status === Office.AsyncResultStatus.Failed) {
//                                    write("Action failed. Error: " + result.error.message);
//                                }
//                                else {
//                                    if (coercionType === Office.CoercionType.Text) {
//                                        dataSelectorSelectedText = result.value;
//                                    }
//                                    else
//                                        if (coercionType === Office.CoercionType.Ooxml) {
//                                            dataSelectorSelectedOOXML = result.value;
//                                            structureOOXML();
//                                        }
//
//                                    if (functionsToExecute.length > 0) {
//                                        // Remove and execute the first function on the queue
//                                        (functionsToExecute.shift())();
//                                    }
//
//                                    //$('#result-text').text(result.value);
//                                }
//                            }
//                        );
//                        // Get the longest word from the selection.
//                        /*var words = range.text.split(/\s+/);
//                        var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });
//        
//                        // Queue a search command.
//                        searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });
//        
//                        // Queue a commmand to load the font property of the results.
//                        context.load(searchResults, 'font');*/
//                    })
//                    .then(context.sync)
//                /*.then(function () {
//                    // Queue a command to highlight the search results.
//                    //searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
//                    //searchResults.items[0].font.bold = true;
//                })
//                .then(context.sync)*/;
//        })
//        .catch(errorHandler);
//}

function dataSelectorSetData(coercionType, dataSource, functionsToExecute) {
    var dataToBeSet = "";

    dataToBeSet += dataSelectorSelectedOOXML.documentBegin;

    if (dataSource === OOXML_SOURCE.TEXT_AREA) {
        //Sets the currentOOXML variable to the current contents of the task pane text area
        dataToBeSet += document.getElementById("dataOOXML").textContent;
    } else {
        if (dataSource === OOXML_SOURCE.MARKER_EDITOR) {
            dataToBeSet +=  dataSelectorSelectedOOXML.textBody;
        }
    }

    dataToBeSet += dataSelectorSelectedOOXML.documentEnd;

    // Check whether we have OOXML in the variable
    if (Office.dataToBeSet !== "" || coercionType === Office.CoercionType.Text) {

        // Call the setSelectedDataAsync, with parameters of:
        // 1. The Data to insert.
        // 2. The coercion type for that data.
        // 3. A callback function that lets us know if it succeeded.

        // Run a batch operation against the Word object model.
        Word.run(function(context) {

                // Create a proxy object for the document body.
                var range = context.document.getSelection();

                // Queue 
                context.load(range, 'text');

                // Synchronize the document state by executing the queued commands
                // and return a promise to indicate task completion.
                return context.sync()
                    .then(function() {

                        // Queue a commmand to insert OOXML in to the beginning of the body.
                        var result;

                        if (coercionType === Office.CoercionType.Ooxml) {
                            result = range.insertOoxml(dataToBeSet, Word.InsertLocation.replace);
                        } else {
                            if (coercionType === Office.CoercionType.Ooxml) {
                                result = range.insertText(dataToBeSet, Word.InsertLocation.replace);
                            }
                        }

                        // Queue 
                        context.load(result, 'text');

                        return context.sync()
                            .then(function() {

                                if (functionsToExecute.length > 0) {
                                    // Remove and execute the first function on the queue
                                    (functionsToExecute.shift())();
                                }
                            })
                            .catch(errorHandler);
                    })
                    .catch(errorHandler);

            })
            .catch(errorHandler);
    }
}

//function _dataSelectorSetData(coercionType, functionsToExecute, alternativeOOXML) {
//    var currentOOXML;
//
//    if (alternativeOOXML === null) {
//        //Sets the currentOOXML variable to the current contents of the task pane text area
//        //        var textArea = document.getElementById("dataOOXML");
//        currentOOXML = document.getElementById("dataOOXML").textContent;
//    } else {
//        currentOOXML = dataSelectorSelectedOOXML.documentBegin +
//            dataSelectorSelectedOOXML.textBody +
//            dataSelectorSelectedOOXML.documentEnd;
//    }
//
//    // Remove all nodes from the status Div so we have a clean space to write to
//    // while (report.hasChildNodes()) {
//    //   report.removeChild(report.lastChild);
//    //}
//
//    // Check whether we have OOXML in the variable
//    if (currentOOXML != "") {
//
//        // Call the setSelectedDataAsync, with parameters of:
//        // 1. The Data to insert.
//        // 2. The coercion type for that data.
//        // 3. A callback function that lets us know if it succeeded.
//
//
//        Office.context.document.setSelectedDataAsync(
//            currentOOXML, { coercionType: "ooxml" },
//            function (result) {
//                // Tell the user we succeeded and then clear the message after a 2 second delay
//                if (result.status == "succeeded") {
//                    //                    report.innerText = "The setOOXML function succeeded!";
//                    setTimeout(function () {
//                        //                        report.innerText = "";
//                    }, 2000);
//
//                    if (functionsToExecute.length > 0) {
//                        // Remove and execute the first function on the queue
//                        (functionsToExecute.shift())();
//                    }
//                }
//                else {
//                    // This runs if the getSliceAsync method does not return a success flag
//                    //                    report.innerText = result.error.message;
//
//                    // Clear the text area just so we don't give you the impression that there's
//                    // valid OOXML waiting to be inserted... 
//                    //                    while (textArea.hasChildNodes()) {
//                    //                        textArea.removeChild(textArea.lastChild);
//                    //                    }
//                }
//            });
//
//    }
//    else {
//
//        // If currentOOXML == "" then we should not even try to insert it, because
//        // that is gauranteed to cause an exception, needlessly.
//        //  report.innerText = "There is currently no OOXML to insert!"
//        //    + " Please select some of your document and click [Get OOXML] first!";
//    }
//}
//

