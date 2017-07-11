function setDeonticMarker(functionsToExecute) {
    setUnsetLogicMarker("<w:color w:val=\"" + "magenta" + "\"/>", "deontic", functionsToExecute);
}

function setTemporalMarker(functionsToExecute) {
    setUnsetLogicMarker("<w:highlight w:val=\"" + "green" + "\"/>", "temporal", functionsToExecute);
}

function setOperationalMarker(functionsToExecute) {
    setUnsetLogicMarker("<w:bdr w:val=\"single\" w:sz=\"4\" w:space=\"2\" w:color=\"" + "cyan" + "\"/>" , "operational", functionsToExecute);
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