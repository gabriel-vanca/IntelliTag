const MARKER_COLOUR = {
    MAGENTA : {
        Text: "magenta",
        Hex: "FF00FF"
},
    GREEN: {
        Text: "green",
        Hex: "00FF00"
    },
    CYAN: {
        Text: "cyan",
        Hex: "00FFFF"
    },
};

function setDeonticMarker(functionsToExecute) {
    setUnsetLogicMarker("<w:color w:val=\"" + MARKER_COLOUR.MAGENTA.Hex + "\"/>", "deontic", functionsToExecute);
}

function setTemporalMarker(functionsToExecute) {
    setUnsetLogicMarker("<w:highlight w:val=\"" + MARKER_COLOUR.GREEN.Text + "\"/>", "temporal", functionsToExecute);
}

function setOperationalMarker(functionsToExecute) {
    setUnsetLogicMarker("<w:bdr w:val=\"single\" w:sz=\"4\" w:space=\"2\" w:color=\"" + MARKER_COLOUR.CYAN.Hex + "\"/>" , "operational", functionsToExecute);
}

function setUnsetLogic(setUnsetLogicFunction) {
    var functionsToExecute = [];
    try{
    functionsToExecute.push(function () { dataSelectorGetOOXML(functionsToExecute); });
    functionsToExecute.push(function () { setUnsetLogicFunction(functionsToExecute); });

//    if (window.dataSelectorSelectedOOXML && window.dataSelectorSelectedOOXML.textBody) {
//        functionsToExecute.push(function() { dataSelectorGetText(functionsToExecute); });
//        functionsToExecute.push(function() { window.setTextArea(window.dataSelectorSelectedOOXML.textBody); });
//    }

    dataSelectorGetText(functionsToExecute);
    } catch (error) {
        errorHandler(error);
    }
}

function isMarkerIsPresent(tag) {
    const stringToLookFor = "IntelliTag_" + tag + "_";
    return dataSelectorSelectedOOXML.textBody.indexOf(stringToLookFor) !== -1;
}

function setUnsetLogicMarker(colour, tag, functionsToExecute) {
    buildGraph();

    if (isMarkerIsPresent(tag)) {
        unmarkText(Graph, colour, tag);
    } else {
        markText(Graph, colour, tag);
    }

    getOOXMLFromGraph();

    dataSelectorSetOOXML(OOXML_SOURCE.MARKER_EDITOR, []);

    if (functionsToExecute.length > 0) {
        // Remove and execute the first function on the queue
        functionsToExecute.shift()();
    }
}