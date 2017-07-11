function setDeonticMarker(functionsToExecute) {
    setLogicMarker("magenta", "deontic", functionsToExecute);
}

function setTemporalMarker(functionsToExecute) {
    setLogicMarker("green", "temporal", functionsToExecute);
}

function setOperationalMarker(functionsToExecute) {
    setLogicMarker("cyan", "operational", functionsToExecute);
}

function setLogic(setLogicFunction) {
    var functionsToExecute = [];
    functionsToExecute.push(function () { dataSelectorGetOOXML(functionsToExecute); });
    functionsToExecute.push(function () { setLogicFunction(functionsToExecute); });

    if (dataSelectorSelectedOOXML && dataSelectorSelectedOOXML.textBody) {
        functionsToExecute.push(function() { dataSelectorGetText(functionsToExecute); });
        functionsToExecute.push(function() { window.setTextArea(dataSelectorSelectedOOXML.textBody); });
    }

    dataSelectorGetText(functionsToExecute);
}

function setLogicMarker(colour, tag, functionsToExecute) {

    buildGraph();
    markText(Graph, colour, tag);
    getOOXMLFromGraph();

    dataSelectorSetOOXML(OOXML_SOURCE.MARKER_EDITOR, []);

    if (functionsToExecute.length > 0) {
        // Remove and execute the first function on the queue
        functionsToExecute.shift()();
    }

  
}