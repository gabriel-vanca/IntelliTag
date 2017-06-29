function setDeonticMarker(functionsToExecute) {

    buildGraph();
    markText(Graph, "FF0000", "deontic");
    getOOXMLFromGraph();

    console.log(dataSelectorSelectedOOXML.textBody);

    if (functionsToExecute.length > 0) {
        // Remove and execute the first function on the queue
        (functionsToExecute.shift())();
    }
     
    dataSelectorSetManualOOXML([]);

    
    //setTextArea(dataSelectorSelectedOOXML.textBody);

    /* var text = dataSelectorSelectedText;
     var ooxml = dataSelectorSelectedOOXML;
 
     var words = text.split(" ");
     for (i = 0; i < words.length; i++) {
         var index = ooxml.indexOf(words[i]);
         if (index == -1)
             continue;
         var finalOoxml = ooxml.substring(0, index);
         finalOoxml += "</w:t> </w: r> <w:r> < w:rPr > <w:color w:val=\"FF0000\" /> </w: rPr > <w:t>";
         finalOoxml += words[i];
         finalOoxml += "</w:t> </w: r> <w:r> <w:t>";
 
     }*/
}