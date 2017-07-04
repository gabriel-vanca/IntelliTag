var Graph;

const INSERT_LOCATION = {
    Begin: 'begin',
    End: 'end',
    Replace: 'replace',
    Intermediary: 'intermediary'
}

function buildGraph() {
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

function createNode(_openTag, _closeTag, _ascendentNode, _textValue, _insertPosition) {

    var node = {
        ascendentNode: _ascendentNode,
        openTag: _openTag,
        closeTag: _closeTag,
        textValue: _textValue,
        listOfDescendentNodes: []
    }

    if (_ascendentNode) {
        if (_insertPosition === INSERT_LOCATION.Intermediary) {
            node.listOfDescendentNodes = _ascendentNode.listOfDescendentNodes;
            _ascendentNode.listOfDescendentNodes = [];
            _ascendentNode.listOfDescendentNodes.push(node);
            node.ascendentNode = _ascendentNode;
            for (let index = 0; index < node.listOfDescendentNodes.length; index++) {
                node.listOfDescendentNodes[index].ascendentNode = node;
            }
        } else if (_insertPosition === INSERT_LOCATION.Replace) {
            _ascendentNode.listOfDescendentNodes = [];
            _ascendentNode.listOfDescendentNodes.push(node);
        } else if (_insertPosition === INSERT_LOCATION.Begin) {
            _ascendentNode.listOfDescendentNodes.unshift(node);
        } else {
            _ascendentNode.listOfDescendentNodes.push(node);
        }
    }

    return node;
}

//function moveAllDescendantNodesToIndermediaryNode(currentNode, intermediaryNode) {
//    intermediaryNode.listOfDescendentNodes = currentNode.listOfDescendentNodes;
//    currentNode.listOfDescendentNodes = [];
//    currentNode.listOfDescendentNodes.push(intermediaryNode);
//    intermediaryNode.ascendentNode = currentNode;
//}

function markText(currentNode, colour, tag) {
    var node;
    if (currentNode.openTag != null &&
        (currentNode.openTag.indexOf("<w:p>") !== -1 || currentNode.openTag.indexOf("<w:p ") !== -1)) {
        let random = Math.floor(Math.random() * 999999999);
        node = createNode("<w:bookmarkStart w:id=\"" + random + "\" w:name=\"" + tag + "_" + random + "\"/>",
            "<w:bookmarkEnd w:id=\"" + random + "\"/>",
            currentNode,
            null,
            INSERT_LOCATION.Intermediary);
    } else if (currentNode.openTag != null &&
        (currentNode.openTag.indexOf("<w:r>") !== -1 || currentNode.openTag.indexOf("<w:r ") !== -1)) {
        var indexOfPropertyTag = -1;
        for (let index = 0; index < currentNode.listOfDescendentNodes.length; index++) {
            node = currentNode.listOfDescendentNodes[index];
            if (node.openTag.indexOf("<w:rPr>") !== -1 || node.openTag.indexOf("<w:rPr ") !== -1) {
                indexOfPropertyTag = index;
            }

            //put here

        }
        // only inexistent case is made
        if (indexOfPropertyTag === -1) {
            var newNode1 = createNode("<w:rPr>", "</w:rPr>", currentNode, null, INSERT_LOCATION.Begin);
            createNode("<w:color w:val=\"" + colour + "\"/>", null, newNode1, null);
        } else {
            var propertyTagNode = currentNode.listOfDescendentNodes[indexOfPropertyTag];
            var freeToAdd = true;

            for (let index = 0; index < propertyTagNode.listOfDescendentNodes.length; index++) {
                if (propertyTagNode.listOfDescendentNodes[index].openTag.indexOf("<w:color") !== -1) {
                    freeToAdd = false;
                    break;
                }
            }

            if (freeToAdd === true) {
                createNode("<w:color w:val=\"" + colour + "\"/>", null, propertyTagNode, null);
            } else {
                freeToAdd = true;

                for (let index = 0; index < propertyTagNode.listOfDescendentNodes.length; index++) {
                    if (propertyTagNode.listOfDescendentNodes[index].openTag.indexOf("<w:highlight") !== -1) {
                        freeToAdd = false;
                        break;
                    }
                }

                if (freeToAdd === true) {
                    createNode("<w:highlight w:val=\"" + colour + "\"/>", null, propertyTagNode, null);
                } else {
                    //freeToAdd = true;
                    createNode("<w:bdr w:val=\"single\" w:sz=\"4\" w:space=\"1\" w:color=\"" + colour + "\"/>",
                        null,
                        propertyTagNode,
                        null);
                }
            }
        }
        return;
    }
    for (let index = 0; index < currentNode.listOfDescendentNodes.length; index++) {
        node = currentNode.listOfDescendentNodes[index];
        markText(node, colour, tag);
    }
}