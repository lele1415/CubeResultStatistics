function createXmlHttp() { 
    if (window.XMLHttpRequest) { 
        xmlHttp = new XMLHttpRequest();
    } else { 
        xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
    } 
} 

function getSourceJs(url) { 
    createXmlHttp();
    xmlHttp.onreadystatechange = writeSource;
    xmlHttp.open("GET", url, true); 
    xmlHttp.send(null); 
} 

function getNextSourceJs(url) {
    xmlHttp.onreadystatechange = writeSource;
    xmlHttp.open("GET", url, true); 
    xmlHttp.send(null);
}

function writeSource() { 
    if (xmlHttp.readyState == 4) { 
        receiveCode(xmlHttp.responseText);
    } 
} 

function option_creat(optionValue, optionInnerHTML)
{
    var option = document.createElement("option");
    option.value = optionValue;
    option.innerHTML = optionInnerHTML;

    return option;
}

function parentNode_appendChild(parentNodeId, node)
{
    var parentNode = document.getElementById(parentNodeId);
    parentNode.appendChild(node);
}

function addOption(SelectId, OptionName, iSeq) 
{
    var option = option_creat(OptionName, OptionName);
    option.value = iSeq;
    parentNode_appendChild(SelectId, option);
}

function removeOption(SelectId, OptionValue)
{
    var parentNode = document.getElementById(SelectId);
    var lenOfChild = parentNode.childNodes.length;
    if (lenOfChild > 0){
        for(var i=0;i<lenOfChild;i++){
            if (parentNode.childNodes[i].value == OptionValue) {
                parentNode.removeChild(parentNode.childNodes[i]);
                break;
            }
        }   
    }
}