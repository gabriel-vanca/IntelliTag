function setDeonticMarker(functionsToExecute) {
    setUnsetLogicMarker("magenta", "deontic", functionsToExecute);
}

function setTemporalMarker(functionsToExecute) {
    setUnsetLogicMarker("green", "temporal", functionsToExecute);
}

function setOperationalMarker(functionsToExecute) {
    setUnsetLogicMarker("cyan", "operational", functionsToExecute);
}

function setUnsetLogic(setUnsetLogicFunction) {
    var functionsToExecute = [];
    functionsToExecute.push(function () { dataSelectorGetOOXML(functionsToExecute); });
    functionsToExecute.push(function () { setUnsetLogicFunction(functionsToExecute); });

//    if (window.dataSelectorSelectedOOXML && window.dataSelectorSelectedOOXML.textBody) {
//        functionsToExecute.push(function() { dataSelectorGetText(functionsToExecute); });
//        functionsToExecute.push(function() { window.setTextArea(window.dataSelectorSelectedOOXML.textBody); });
//    }

    dataSelectorGetText(functionsToExecute);
}

function checkIfMarkerIsPresent(tag) {
    const stringToLookFor = "IntelliTag_" + tag + "_";
    return (dataSelectorSelectedOOXML.textBody.indexOf(stringToLookFor) === -1);
}

function setUnsetLogicMarker(colour, tag, functionsToExecute) {

    buildGraph();

    markText(Graph, colour, tag);
    getOOXMLFromGraph();

    dataSelectorSetOOXML(OOXML_SOURCE.MARKER_EDITOR, []);

    if (functionsToExecute.length > 0) {
        // Remove and execute the first function on the queue
        functionsToExecute.shift()();
    }

  
}