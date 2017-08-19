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
