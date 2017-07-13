function structureOOXML() {

    var indexBegin, indexEnd;

//    window.dataSelectorSelectedOOXML = { documentBegin: "", textBody: dataSelectorSelectedOOXML, documentEnd: "" };
    try{
    indexBegin = window.dataSelectorSelectedOOXML.indexOf("<w:body");
    for (; indexBegin < window.dataSelectorSelectedOOXML.length; indexBegin++) {
        if (window.dataSelectorSelectedOOXML.charAt(indexBegin) === ">") {
            indexBegin++;
            break;
        }
    }
    indexEnd = window.dataSelectorSelectedOOXML.indexOf("</w:body>");
    const documentBegin = window.dataSelectorSelectedOOXML.substring(0, indexBegin);
    const documentEnd = window.dataSelectorSelectedOOXML.substring(indexEnd, window.dataSelectorSelectedOOXML.length);
    const textBody = window.dataSelectorSelectedOOXML.substring(indexBegin, indexEnd);
    window.dataSelectorSelectedOOXML = { documentBegin: documentBegin, textBody: textBody, documentEnd: documentEnd };
    } catch (error) {
        errorHandler(error);
    }
}

function copyString(_string) {
    var newString = "";
    try{
    for (var i = 0; i < _string.length; i++)
            newString += _string.charAt(i);
    } catch (error) {
        errorHandler(error);
    }

    return newString;
}