function setDeonticMarker() {
    var text = dataSelectorSelectedText;
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

    }