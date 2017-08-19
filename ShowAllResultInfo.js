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

function parentNode_appendChild(parentNodeId, node)
{
    var parentNode = document.getElementById(parentNodeId);
    parentNode.appendChild(node);
}

function option_creat(optionValue, optionInnerHTML)
{
    var option = document.createElement("option");
    option.value = optionValue;
    option.innerHTML = optionInnerHTML;

    return option;
}

function setInnerHtmlById(id, text) {
    document.getElementById(id).innerHTML = text;
}

function setValueById(id, value) {
    document.getElementById(id).value = value;
}

function getValueById(id) {
    return document.getElementById(id).value;
}