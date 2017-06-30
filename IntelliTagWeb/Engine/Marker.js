function setDeonticMarker(functionsToExecute) {
    setLogicMarker("magenta", "deontic", functionsToExecute);
}

function setTemporalMarker(functionsToExecute) {
    setLogicMarker("green", "deontic", functionsToExecute);
}

function setOperationalMarker(functionsToExecute) {
    setLogicMarker("cyan", "deontic", functionsToExecute);
}

function setLogicMarker(colour, tag, functionsToExecute) {

    buildGraph();
    markText(Graph, colour, tag);
    getOOXMLFromGraph();

//    console.log(dataSelectorSelectedOOXML.textBody);

    if (functionsToExecute.length > 0) {
        // Remove and execute the first function on the queue
        (functionsToExecute.shift())();
    }

    dataSelectorSetOOXML(OOXML_SOURCE.MARKER_EDITOR, []);
}