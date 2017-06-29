var Graph;

function buildGraph(string) {
    Graph = createNode(null, null, null, null);
    var ascendentNode = Graph;
    var tempString = copyString(dataSelectorSelectedOOXML.textBody);
 
    while (true) {
        if (tempString == null || tempString.length < 1)
            break;
        if (tempString[tempString.length - 1] !== " ") {
            tempString += "  ";
        }
        var simpleTag = false;
        var closeTag = false;

        var indexBegin = tempString.indexOf("<");
        if (indexBegin === -1)
            break;
        var indexEnd = tempString.indexOf(">");
        if (indexEnd === -1)
            break;

        
        for (let index = indexEnd-1; index > indexBegin; index--) {
            if (tempString.charAt(index) === " ")
                continue;

            if (tempString.charAt(index) === "/") {
                simpleTag = true;
                break;
            }
            break;
        }

        if (simpleTag === false) {
            for (let index = indexBegin+1; index < indexEnd; index++) {
                if (tempString.charAt(index) === " ")
                    continue;

                if (tempString.charAt(index) === "/") {
                    closeTag = true;
                    break;
                }
                break;
            }
        }

        var tag = tempString.substring(indexBegin, indexEnd + 1);
        tempString = tempString.substring(indexEnd + 1, tempString.length);

        if (simpleTag === true) {
            createNode(tag, null, ascendentNode, null);
            
        } else {
            if (closeTag === false) {
                var node = createNode(tag, null, ascendentNode, null);
                ascendentNode = node;
               
                if (tag.indexOf("<w:t") !== -1) {
                    var indexTextBegin = 0;
                    var indexTextEnd = tempString.indexOf("</");
                    if (indexTextEnd !== -1) {
                        var textValue = tempString.substring(indexTextBegin, indexTextEnd);
                        createNode(null, null, ascendentNode, textValue);
                        tempString = tempString.substring(indexTextEnd, tempString.length);
                    }
                }
            } else {
                //if (ascendentNode) {
                ascendentNode.closeTag = tag;
                ascendentNode = ascendentNode.ascendentNode;
                //}
            }
        }
    }
}

function constructOOXMLFromGraph(currentNode) {
    if (currentNode.openTag != null)
        window.dataSelectorSelectedOOXML.textBody += currentNode.openTag;
    if (currentNode.textValue != null)
        window.dataSelectorSelectedOOXML.textBody += currentNode.textValue;

    for (let index = 0; index < currentNode.listOfDescendentNodes.length; index++) {
        var node = currentNode.listOfDescendentNodes[index];
        constructOOXMLFromGraph(node);
    }

    if (currentNode.closeTag != null)
        window.dataSelectorSelectedOOXML.textBody += currentNode.closeTag;
}

function getOOXMLFromGraph() {
    window.dataSelectorSelectedOOXML.textBody = "";
    constructOOXMLFromGraph(Graph);
}


//function structureOOXML() {
//    const indexBegin = window.dataSelectorSelectedOOXML.indexOf("<w:body");
//    const indexEnd = window.dataSelectorSelectedOOXML.indexOf("</w:body>") + 9;
//    const documentBegin = window.dataSelectorSelectedOOXML.substring(0, indexBegin);
//    const documentEnd = window.dataSelectorSelectedOOXML.substring(indexEnd, window.dataSelectorSelectedOOXML.length);
//    const textBody = window.dataSelectorSelectedOOXML.substring(indexBegin, indexEnd);
//    window.dataSelectorSelectedOOXML = { documentBegin: documentBegin, textBody: textBody, documentEnd: documentEnd };
//}

function createNode(_openTag, _closeTag, _ascendentNode, _textValue) {

    var node = {
        ascendentNode: _ascendentNode,
        openTag: _openTag,
        closeTag: _closeTag,
        textValue: _textValue,
        listOfDescendentNodes: []
    }

    if (_ascendentNode) {
        _ascendentNode.listOfDescendentNodes.push(node);
    }

    return node;
}

function markText(currentNode, colour, tag) {
    var node;
    if (currentNode.openTag != null && (currentNode.openTag.indexOf("<w:r>") !== -1 || currentNode.openTag.indexOf("<w:r ") !== -1)) {
        var indexOfPropertyTag = -1;
        for (let index = 0; index < currentNode.listOfDescendentNodes.length; index++) {
            node = currentNode.listOfDescendentNodes[index];
            if (node.openTag.indexOf("<w:rPr>") != -1)
                indexOfPropertyTag = index;
        }
        // only inexistent case is made
        if (indexOfPropertyTag === -1) {
            var newNode1 = createNode("<w:rPr>", "</w:rPr>", currentNode, null);
//            currentNode.listOfDescendentNodes.unshift(newNode1);
            createNode("<w:color w:val=\"" + colour + "\"/>", null, newNode1, null);
        }

    } else {
        for (let index = 0; index < currentNode.listOfDescendentNodes.length; index++) {
            node = currentNode.listOfDescendentNodes[index];
            markText(node, colour, tag);
        }
    }
}