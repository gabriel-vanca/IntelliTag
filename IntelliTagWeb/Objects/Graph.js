var Graph;

const INSERT_LOCATION = {
    Begin: 'begin',
    End: 'end',
    Replace: 'replace',
    Intermediary: 'intermediary'
};

function buildGraph() {
    Graph = createNode(null, null, null, null);
    var ascendentNode = Graph;
    var tempString = copyString(dataSelectorSelectedOOXML.textBody);
 
    for (; ;) {
        if (tempString === null || typeof tempString === "undefined" || tempString.length < 1)
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
    if (currentNode.openTag !== null && typeof currentNode.openTag !== "undefined")
        window.dataSelectorSelectedOOXML.textBody += currentNode.openTag;
    if (currentNode.textValue !== null && typeof currentNode.textValue !== "undefined" && currentNode.textValue !== "")
        window.dataSelectorSelectedOOXML.textBody += currentNode.textValue;

    for (let index = 0; index < currentNode.listOfDescendentNodes.length; index++) {
        const node = currentNode.listOfDescendentNodes[index];
        constructOOXMLFromGraph(node);
    }

    if (currentNode.closeTag !== null && typeof currentNode.closeTag !== "undefined")
        window.dataSelectorSelectedOOXML.textBody += currentNode.closeTag;
}

function getOOXMLFromGraph() {
    window.dataSelectorSelectedOOXML.textBody = "";
    constructOOXMLFromGraph(Graph);
}

function createNode(_openTag, _closeTag, _ascendentNode, _textValue, _insertPosition) {

    const node = {
        ascendentNode: _ascendentNode,
        openTag: _openTag,
        closeTag: _closeTag,
        textValue: _textValue,
        listOfDescendentNodes: []
    };

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

function removeNode(_nodeToBeRemoved) {
    const _ascentNode = _nodeToBeRemoved.ascendentNode;
    var position = -1;
    for (let index = 0; index < _ascentNode.listOfDescendentNodes.length; index++) {
        if (_ascentNode.listOfDescendentNodes[index] === _nodeToBeRemoved) {
            position = index;
            break;
        }
    }
    if (position === -1)
        return;

    if (_nodeToBeRemoved.listOfDescendentNodes !== null && typeof _nodeToBeRemoved.listOfDescendentNodes !== "undefined" && _nodeToBeRemoved.listOfDescendentNodes.length > 0)
    {

    for (let index = 0; index < _nodeToBeRemoved.listOfDescendentNodes.length; index++) {
        _ascentNode.listOfDescendentNodes.splice(index + 1, 0, _nodeToBeRemoved.listOfDescendentNodes[index]);
        }
    }

    _ascentNode.listOfDescendentNodes.splice(position, 1);
    _nodeToBeRemoved.openTag = null;
}

function markText(currentNode, visibleTag, logicTag) {
    var node;
    if (currentNode.openTag !== null && typeof currentNode.openTag !== "undefined" &&
        (currentNode.openTag.indexOf("<w:p>") !== -1 || currentNode.openTag.indexOf("<w:p ") !== -1)) {
        let random = Math.floor(Math.random() * 999999999);
        createNode("<w:bookmarkStart w:id=\"" +
            random +
            "\" w:name=\"" +
            "IntelliTag_"+
            logicTag +
            "_" +
            Settings.lastLogicId +
            "_" +
            random +
            "\"/>",
            "<w:bookmarkEnd w:id=\"" + random + "\"/>",
            currentNode,
            null,
            INSERT_LOCATION.Intermediary);
        Settings.lastLogicId++;
    } else if (currentNode.openTag !== null && typeof currentNode.openTag !== "undefined" &&
        (currentNode.openTag.indexOf("<w:r>") !== -1 || currentNode.openTag.indexOf("<w:r ") !== -1)) {
        var indexOfPropertyTag = -1;
        for (let index = 0; index < currentNode.listOfDescendentNodes.length; index++) {
            node = currentNode.listOfDescendentNodes[index];
            if (node.openTag.indexOf("<w:rPr>") !== -1 || node.openTag.indexOf("<w:rPr ") !== -1) {
                indexOfPropertyTag = index;
                break;
            }
        }
        if (indexOfPropertyTag === -1) {
            var newNode1 = createNode("<w:rPr>", "</w:rPr>", currentNode, null, INSERT_LOCATION.Begin);
            createNode(visibleTag, null, newNode1, null);
        } else {
            var propertyTagNode = currentNode.listOfDescendentNodes[indexOfPropertyTag];
            createNode(visibleTag, null, propertyTagNode, null);

        }
        return;
    }
    for (let index = 0; index < currentNode.listOfDescendentNodes.length; index++) {
        node = currentNode.listOfDescendentNodes[index];
        markText(node, visibleTag, logicTag);
    }
}

function unmarkText(currentNode, visibleTag, logicTag) {
    var node;
    if (currentNode.openTag !== null &&
        typeof currentNode.openTag !== "undefined" &&
        currentNode.openTag.indexOf("<w:bookmarkStart") !== -1 &&
        currentNode.openTag.indexOf("IntelliTag_" + logicTag + "_") !== -1) {
        removeNode(currentNode);
//        currentNode = null;
//        return;
    } else if (currentNode.openTag !== null &&
        typeof currentNode.openTag !== "undefined" &&
        currentNode.openTag.indexOf(visibleTag) !== -1) {
        removeNode(currentNode);
        return;
    }
    for (let index = 0; index < currentNode.listOfDescendentNodes.length; index++) {
        node = currentNode.listOfDescendentNodes[index];
        unmarkText(node, visibleTag, logicTag);

        if ((node.openTag === null || typeof node.openTag === "undefined" || node.openTag === []) && (node.textBody !== null && typeof node.textBody !== "undefined" && node.textBody.length > 0))
            index--;
    }
}