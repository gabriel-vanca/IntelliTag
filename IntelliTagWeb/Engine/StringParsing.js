function structureOOXML() {

    window.dataSelectorSelectedOOXML = { documentBegin: "", textBody: dataSelectorSelectedOOXML, documentEnd: "" };

//    var indexBegin = window.dataSelectorSelectedOOXML.indexOf("<w:body");
//    for (; indexBegin < window.dataSelectorSelectedOOXML.length; indexBegin++) {
//        if (window.dataSelectorSelectedOOXML.charAt(indexBegin) == ">") {
//            indexBegin++;
//            break;
//        }
//    }
//   // var textBody = window.dataSelectorSelectedOOXML.substring(indexBegin, indexEnd);
//    const indexEnd = window.dataSelectorSelectedOOXML.indexOf("</w:body>");
//    const documentBegin = "";// window.dataSelectorSelectedOOXML.substring(0, indexBegin);
//    const documentEnd = ""; //window.dataSelectorSelectedOOXML.substring(indexEnd, window.dataSelectorSelectedOOXML.length);
//    const textBody = window.dataSelectorSelectedOOXML.substring(indexBegin, indexEnd);
//    window.dataSelectorSelectedOOXML = { documentBegin: documentBegin, textBody: textBody, documentEnd: documentEnd };

}

function copyString(_string) {
    var newString = "";
    for (var i = 0; i < _string.length; i++)
        newString += _string.charAt(i);

    return newString;
}

function lookForNextTag(_string) {

    var singleLine = false;
    var indexBeginBegin = _string.indexOf("<");
    var indexBeginEnd = _string.indexOf(">");


    for (var index = indexBeginEnd; index > indexBeginBegin; i--) {
        if (_string.charAt(index) == " ")
            continue;

        if (_string.charAt(index) == "/") {
            singleLine = true;
            break;
        }
        break;
    }

    if (singleLine == true) {

    } else {
        var indexEndBegin = _string.lastIndexOf("</");
        var indexEndEnd = _string.lastIndexOf(">");
    }


}