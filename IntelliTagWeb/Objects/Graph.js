function buildGraph(string) {
    var ascendentNode = null;


}

function createNode(_openTag, _closeTag, _ascendentNode, _textValue) {

    var node = {
        ascendentNode: _ascendentNode,
        openTag: _openTag,
        closeTag: _closeTag,
        textValue: _textValue,
        listOfDescendentNodes: []
    }

    if (previousNode) {
        previousNode.listOfDescendentNodes.push(node);
    }
}