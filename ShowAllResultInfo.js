var ID_VALID_RESULTS = "valid_results"
var ID_INVALID_RESULTS = "invalid_results"
var ID_BR_RESULTS = "br_results"

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

                if (SelectId != ID_VALID_RESULTS) {
                    if (i < lenOfChild - 1) {
                        parentNode.childNodes[i].selected = true;
                        showResultInfo(SelectId);
                    } else if (i > 0) {
                        parentNode.childNodes[i - 1].selected = true;
                        showResultInfo(SelectId);
                    } else {
                        clearAllResultInfoText();
                    }
                }
                
                break;
            }
        }   
    }
}

function selectAnotherOption(SelectId, OptionValue, which)
{
    var parentNode = document.getElementById(SelectId);
    var lenOfChild = parentNode.childNodes.length;
    if (lenOfChild > 0){
        for(var i=0;i<lenOfChild;i++){
            if (parentNode.childNodes[i].value == OptionValue) {
                parentNode.childNodes[i].selected = false;

                if (which == 0 && i < lenOfChild - 1) {
                    parentNode.childNodes[i + 1].selected = true;
                    showResultInfo(SelectId);
                } else if (which == 1 && i > 0) {
                    parentNode.childNodes[i - 1].selected = true;
                    showResultInfo(SelectId);
                }
                break;
            }
        }   
    }
}

function clearAllSelectedOption(exceptSelectId) {
    if (exceptSelectId != ID_VALID_RESULTS) {
        clearSelectedOption(ID_VALID_RESULTS);
    }
    if (exceptSelectId != ID_INVALID_RESULTS) {
        clearSelectedOption(ID_INVALID_RESULTS);
    }
    if (exceptSelectId != ID_BR_RESULTS) {
        clearSelectedOption(ID_BR_RESULTS);
    }
}

function clearSelectedOption(selectId) {
    var parentNode = document.getElementById(selectId);
    var index = parentNode.selectedIndex;
    if (index >= 0) {
        parentNode.childNodes[index].selected = false;
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