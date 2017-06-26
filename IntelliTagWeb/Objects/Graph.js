//require("./Node.js");

var Graph = [];

function buildGraph(string) {
    var previousNode = null;
    Graph = null;
}

function createNode(openTag, closeTag, previousNode, text) {
    if (previousNode) {
        previousNode.listOfNextNodes.push(this);
    }
    this.previousNode = previousNode;
    this.openTag = openTag;
    this.closeTag = closeTag;
    this.text = text;
}