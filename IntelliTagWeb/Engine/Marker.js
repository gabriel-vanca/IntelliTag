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
//    functionsToExecute.push(function () { window.setTextArea(dataSelectorSelectedOOXML.textBody); });

    dataSelectorGetText(functionsToExecute);
}

function setLogicMarker(colour, tag, functionsToExecute) {

    buildGraph();
    markText(Graph, colour, tag);
    getOOXMLFromGraph();

    if (functionsToExecute.length > 0) {
        // Remove and execute the first function on the queue
        (functionsToExecute.shift())();
    }

    dataSelectorSetOOXML(OOXML_SOURCE.MARKER_EDITOR, []);
}