"use strict";

importScripts(
    './jszip.min.js',
    './highlight.min.js',
    './colz.class.min.js',
    './highlight.min.js',
    './tXml.js',
    './functions.js'
);

var MsgQueue = new Array();

var themeContent = null;

var chartID = 0;

var titleFontSize = 42;
var bodyFontSize = 20;
var otherFontSize = 16;

var styleTable = {};

onmessage = function(e) {
    
    switch (e.data.type) {
        case "processPPTX":
            processPPTX(e.data.data);
            break;
        case "getMsgQueue":
            self.postMessage({
                "type": "processMsgQueue",
                "data": MsgQueue
            });
            break;
        default:
    }

}

function processPPTX(data) {
    
    var dateBefore = new Date();
    
    var zip = new JSZip(data);
    
    if (zip.file("docProps/thumbnail.jpeg") !== null) {
        var pptxThumbImg = base64ArrayBuffer(zip.file("docProps/thumbnail.jpeg").asArrayBuffer());
        self.postMessage({
            "type": "pptx-thumb",
            "data": pptxThumbImg
        });
    }
    
    var filesInfo = getContentTypes(zip);
    var slideSize = getSlideSize(zip);
    themeContent = loadTheme(zip);
    
    self.postMessage({
        "type": "slideSize",
        "data": slideSize
    });
    
    var numOfSlides = filesInfo["slides"].length;
    for (var i=0; i<numOfSlides; i++) {
        var filename = filesInfo["slides"][i];
        var slideHtml = processSingleSlide(zip, filename, i, slideSize);
        self.postMessage({
            "type": "slide",
            "data": slideHtml
        });
        self.postMessage({
            "type": "progress-update",
            "data": (i + 1) * 100 / numOfSlides
        });
    }

    self.postMessage({
        "type": "globalCSS",
        "data": genGlobalCSS()
    });
    
    var dateAfter = new Date();
    self.postMessage({
        "type": "ExecutionTime",
        "data": dateAfter - dateBefore
    });
    
}

function readXmlFile(zip, filename) {
    return tXml(zip.file(filename).asText());
}

function getContentTypes(zip) {
    var ContentTypesJson = readXmlFile(zip, "[Content_Types].xml");
    var subObj = ContentTypesJson["Types"]["Override"];
    var slidesLocArray = [];
    var slideLayoutsLocArray = [];
    for (var i=0; i<subObj.length; i++) {
        switch (subObj[i]["attrs"]["ContentType"]) {
            case "application/vnd.openxmlformats-officedocument.presentationml.slide+xml":
                slidesLocArray.push(subObj[i]["attrs"]["PartName"].substr(1));
                break;
            case "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml":
                slideLayoutsLocArray.push(subObj[i]["attrs"]["PartName"].substr(1));
                break;
            default:
        }
    }
    return {
        "slides": slidesLocArray,
        "slideLayouts": slideLayoutsLocArray
    };
}

function getSlideSize(zip) {
    // Pixel = EMUs * Resolution / 914400;  (Resolution = 96)
    var content = readXmlFile(zip, "ppt/presentation.xml");
    var sldSzAttrs = content["p:presentation"]["p:sldSz"]["attrs"]
    return {
        "width": parseInt(sldSzAttrs["cx"]) * 96 / 914400,
        "height": parseInt(sldSzAttrs["cy"]) * 96 / 914400
    };
}

function loadTheme(zip) {
    var preResContent = readXmlFile(zip, "ppt/_rels/presentation.xml.rels");
    var relationshipArray = preResContent["Relationships"]["Relationship"];
    var themeURI = undefined;
    if (relationshipArray.constructor === Array) {
        for (var i=0; i<relationshipArray.length; i++) {
            if (relationshipArray[i]["attrs"]["Type"] === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme") {
                themeURI = relationshipArray[i]["attrs"]["Target"];
                break;
            }
        }
    } else if (relationshipArray["attrs"]["Type"] === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme") {
        themeURI = relationshipArray["attrs"]["Target"];
    }
    
    if (themeURI === undefined) {
        throw Error("Can't open theme file.");
    }
    
    return readXmlFile(zip, "ppt/" + themeURI);
}

function processSingleSlide(zip, sldFileName, index, slideSize) {
    
    self.postMessage({
        "type": "INFO",
        "data": "Processing slide" + (index + 1)
    });
    
    // =====< Step 1 >=====
    // Read relationship filename of the slide (Get slideLayoutXX.xml)
    // @sldFileName: ppt/slides/slide1.xml
    // @resName: ppt/slides/_rels/slide1.xml.rels
    var resName = sldFileName.replace("slides/slide", "slides/_rels/slide") + ".rels";
    var resContent = readXmlFile(zip, resName);
    var RelationshipArray = resContent["Relationships"]["Relationship"];
    var layoutFilename = "";
    var slideResObj = {};
    if (RelationshipArray.constructor === Array) {
        for (var i=0; i<RelationshipArray.length; i++) {
            switch (RelationshipArray[i]["attrs"]["Type"]) {
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout":
                    layoutFilename = RelationshipArray[i]["attrs"]["Target"].replace("../", "ppt/");
                    break;
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide":
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image":
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart":
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink":
                default:
                    slideResObj[RelationshipArray[i]["attrs"]["Id"]] = {
                        "type": RelationshipArray[i]["attrs"]["Type"].replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/", ""),
                        "target": RelationshipArray[i]["attrs"]["Target"].replace("../", "ppt/")
                    };
            }
        }
    } else {
        layoutFilename = RelationshipArray["attrs"]["Target"].replace("../", "ppt/");
    }
    
    // Open slideLayoutXX.xml
    var slideLayoutContent = readXmlFile(zip, layoutFilename);
    var slideLayoutTables = indexNodes(slideLayoutContent);
    //debug(slideLayoutTables);
    
    // =====< Step 2 >=====
    // Read slide master filename of the slidelayout (Get slideMasterXX.xml)
    // @resName: ppt/slideLayouts/slideLayout1.xml
    // @masterName: ppt/slideLayouts/_rels/slideLayout1.xml.rels
    var slideLayoutResFilename = layoutFilename.replace("slideLayouts/slideLayout", "slideLayouts/_rels/slideLayout") + ".rels";
    var slideLayoutResContent = readXmlFile(zip, slideLayoutResFilename);
    RelationshipArray = slideLayoutResContent["Relationships"]["Relationship"];
    var masterFilename = "";
    if (RelationshipArray.constructor === Array) {
        for (var i=0; i<RelationshipArray.length; i++) {
            switch (RelationshipArray[i]["attrs"]["Type"]) {
                case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster":
                    masterFilename = RelationshipArray[i]["attrs"]["Target"].replace("../", "ppt/");
                    break;
                default:
            }
        }
    } else {
        masterFilename = RelationshipArray["attrs"]["Target"].replace("../", "ppt/");
    }
    // Open slideMasterXX.xml
    var slideMasterContent = readXmlFile(zip, masterFilename);
    var slideMasterTextStyles = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:txStyles"]);
    var slideMasterTables = indexNodes(slideMasterContent);
    //debug(slideMasterTables);
    
    
    // =====< Step 3 >=====
    var slideContent = readXmlFile(zip, sldFileName);
    var nodes = slideContent["p:sld"]["p:cSld"]["p:spTree"];
    var warpObj = {
        "zip": zip,
        "slideLayoutTables": slideLayoutTables,
        "slideMasterTables": slideMasterTables,
        "slideResObj": slideResObj,
        "slideMasterTextStyles": slideMasterTextStyles
    };
    
    var bgColor = getSlideBackgroundFill(slideContent, slideLayoutContent, slideMasterContent);
    
    var result = "<section style='width:" + slideSize.width + "px; height:" + slideSize.height + "px; background-color: #" + bgColor + "'>"
    
    for (var nodeKey in nodes) {
        if (nodes[nodeKey].constructor === Array) {
            for (var i=0; i<nodes[nodeKey].length; i++) {
                result += processNodesInSlide(nodeKey, nodes[nodeKey][i], warpObj);
            }
        } else {
            result += processNodesInSlide(nodeKey, nodes[nodeKey], warpObj);
        }
    }
    
    return result + "</section>";
}

function indexNodes(content) {
    
    var keys = Object.keys(content);
    var spTreeNode = content[keys[0]]["p:cSld"]["p:spTree"];
    
    var idTable = {};
    var idxTable = {};
    var typeTable = {};
    
    for (var key in spTreeNode) {

        if (key == "p:nvGrpSpPr" || key == "p:grpSpPr") {
            continue;
        }
        
        var targetNode = spTreeNode[key];
        
        if (targetNode.constructor === Array) {
            for (var i=0; i<targetNode.length; i++) {
                var nvSpPrNode = targetNode[i]["p:nvSpPr"];
                var id = getTextByPathList(nvSpPrNode, ["p:cNvPr", "attrs", "id"]);
                var idx = getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "idx"]);
                var type = getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "type"]);
                
                if (id !== undefined) {
                    idTable[id] = targetNode[i];
                }
                if (idx !== undefined) {
                    idxTable[idx] = targetNode[i];
                }
                if (type !== undefined) {
                    typeTable[type] = targetNode[i];
                }
            }
        } else {
            var nvSpPrNode = targetNode["p:nvSpPr"];
            var id = getTextByPathList(nvSpPrNode, ["p:cNvPr", "attrs", "id"]);
            var idx = getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "idx"]);
            var type = getTextByPathList(nvSpPrNode, ["p:nvPr", "p:ph", "attrs", "type"]);
            
            if (id !== undefined) {
                idTable[id] = targetNode;
            }
            if (idx !== undefined) {
                idxTable[idx] = targetNode;
            }
            if (type !== undefined) {
                typeTable[type] = targetNode;
            }
        }
        
    }
    
    return {"idTable": idTable, "idxTable": idxTable, "typeTable": typeTable};
}

function processNodesInSlide(nodeKey, nodeValue, warpObj) {
    
    var result = "";
    
    switch (nodeKey) {
        case "p:sp":    // Shape, Text
            result = processSpNode(nodeValue, warpObj);
            break;
        case "p:cxnSp":    // Shape, Text (with connection)
            result = processCxnSpNode(nodeValue, warpObj);
            break;
        case "p:pic":    // Picture
            result = processPicNode(nodeValue, warpObj);
            break;
        case "p:graphicFrame":    // Chart, Diagram, Table
            result = processGraphicFrameNode(nodeValue, warpObj);
            break;
        case "p:grpSp":    // 群組
            result = processGroupSpNode(nodeValue, warpObj);
            break;
        default:
    }
    
    return result;
    
}

function processGroupSpNode(node, warpObj) {
    
    var factor = 96 / 914400;
    
    var xfrmNode = node["p:grpSpPr"]["a:xfrm"];
    var x = parseInt(xfrmNode["a:off"]["attrs"]["x"]) * factor;
    var y = parseInt(xfrmNode["a:off"]["attrs"]["y"]) * factor;
    var chx = parseInt(xfrmNode["a:chOff"]["attrs"]["x"]) * factor;
    var chy = parseInt(xfrmNode["a:chOff"]["attrs"]["y"]) * factor;
    var cx = parseInt(xfrmNode["a:ext"]["attrs"]["cx"]) * factor;
    var cy = parseInt(xfrmNode["a:ext"]["attrs"]["cy"]) * factor;
    var chcx = parseInt(xfrmNode["a:chExt"]["attrs"]["cx"]) * factor;
    var chcy = parseInt(xfrmNode["a:chExt"]["attrs"]["cy"]) * factor;
    
    var order = node["attrs"]["order"];
    
    var result = "<div class='block group' style='z-index: " + order + "; top: " + (y - chy) + "px; left: " + (x - chx) + "px; width: " + (cx - chcx) + "px; height: " + (cy - chcy) + "px;'>";
    
    // Procsee all child nodes
    for (var nodeKey in node) {
        if (node[nodeKey].constructor === Array) {
            for (var i=0; i<node[nodeKey].length; i++) {
                result += processNodesInSlide(nodeKey, node[nodeKey][i], warpObj);
            }
        } else {
            result += processNodesInSlide(nodeKey, node[nodeKey], warpObj);
        }
    }
    
    result += "</div>";
    
    return result;
}

function processSpNode(node, warpObj) {
    
    /*
     *  958    <xsd:complexType name="CT_GvmlShape">
     *  959   <xsd:sequence>
     *  960     <xsd:element name="nvSpPr" type="CT_GvmlShapeNonVisual"     minOccurs="1" maxOccurs="1"/>
     *  961     <xsd:element name="spPr"   type="CT_ShapeProperties"        minOccurs="1" maxOccurs="1"/>
     *  962     <xsd:element name="txSp"   type="CT_GvmlTextShape"          minOccurs="0" maxOccurs="1"/>
     *  963     <xsd:element name="style"  type="CT_ShapeStyle"             minOccurs="0" maxOccurs="1"/>
     *  964     <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
     *  965   </xsd:sequence>
     *  966 </xsd:complexType>
     */
    
    var id = node["p:nvSpPr"]["p:cNvPr"]["attrs"]["id"];
    var name = node["p:nvSpPr"]["p:cNvPr"]["attrs"]["name"];
    var idx = (node["p:nvSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["idx"];
    var type = (node["p:nvSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["type"];
    var order = node["attrs"]["order"];
    
    var slideLayoutSpNode = undefined;
    var slideMasterSpNode = undefined;
    
    if (type !== undefined) {
        if (idx !== undefined) {
            slideLayoutSpNode = warpObj["slideLayoutTables"]["typeTable"][type];
            slideMasterSpNode = warpObj["slideMasterTables"]["typeTable"][type];
        } else {
            slideLayoutSpNode = warpObj["slideLayoutTables"]["typeTable"][type];
            slideMasterSpNode = warpObj["slideMasterTables"]["typeTable"][type];
        }
    } else {
        if (idx !== undefined) {
            slideLayoutSpNode = warpObj["slideLayoutTables"]["idxTable"][idx];
            slideMasterSpNode = warpObj["slideMasterTables"]["idxTable"][idx];
        } else {
            // Nothing
        }
    }
    
    if (type === undefined) {
        type = getTextByPathList(slideLayoutSpNode, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
        if (type === undefined) {
            type = getTextByPathList(slideMasterSpNode, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
        }
    }
    
    debug( {"id": id, "name": name, "idx": idx, "type": type, "order": order} );
    //debug( JSON.stringify( node ) );
    
    return genShape(node, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, warpObj);
}

function processCxnSpNode(node, warpObj) {

    var id = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["id"];
    var name = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["name"];
    //var idx = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["idx"];
    //var type = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["type"];
    //<p:cNvCxnSpPr>(<p:cNvCxnSpPr>, <a:endCxn>)
    var order = node["attrs"]["order"];

    debug( {"id": id, "name": name, "order": order} );
    
    return genShape(node, undefined, undefined, id, name, undefined, undefined, order, warpObj);
}

function genShape(node, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order, warpObj) {
    
    var xfrmList = ["p:spPr", "a:xfrm"];
    var slideXfrmNode = getTextByPathList(node, xfrmList);
    var slideLayoutXfrmNode = getTextByPathList(slideLayoutSpNode, xfrmList);
    var slideMasterXfrmNode = getTextByPathList(slideMasterSpNode, xfrmList);
    
    var result = "";
    var shapType = getTextByPathList(node, ["p:spPr", "a:prstGeom", "attrs", "prst"]);
    
    var isFlipV = false;
    if ( getTextByPathList(slideXfrmNode, ["attrs", "flipV"]) === "1" || getTextByPathList(slideXfrmNode, ["attrs", "flipH"]) === "1") {
        isFlipV = true;
    }
    
    if (shapType !== undefined) {
        
        var off = getTextByPathList(slideXfrmNode, ["a:off", "attrs"]);
        var x = parseInt(off["x"]) * 96 / 914400;
        var y = parseInt(off["y"]) * 96 / 914400;
        
        var ext = getTextByPathList(slideXfrmNode, ["a:ext", "attrs"]);
        var w = parseInt(ext["cx"]) * 96 / 914400;
        var h = parseInt(ext["cy"]) * 96 / 914400;
        
        result += "<svg class='drawing' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
                "' style='" + 
                    getPosition(slideXfrmNode, undefined, undefined) + 
                    getSize(slideXfrmNode, undefined, undefined) +
                    " z-index: " + order + ";" +
                "'>";
        
        // Fill Color
        var fillColor = getShapeFill(node, true);
        
        // Border Color        
        var border = getBorder(node, true);
        
        var headEndNodeAttrs = getTextByPathList(node, ["p:spPr", "a:ln", "a:headEnd", "attrs"]);
        var tailEndNodeAttrs = getTextByPathList(node, ["p:spPr", "a:ln", "a:tailEnd", "attrs"]);
        // type: none, triangle, stealth, diamond, oval, arrow
        if ( (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) || 
             (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) ) {
            var triangleMarker = "<defs><marker id=\"markerTriangle\" viewBox=\"0 0 10 10\" refX=\"1\" refY=\"5\" markerWidth=\"5\" markerHeight=\"5\" orient=\"auto-start-reverse\" markerUnits=\"strokeWidth\"><path d=\"M 0 0 L 10 5 L 0 10 z\" /></marker></defs>";
            result += triangleMarker;
        }
        
        switch (shapType) {
            case "accentBorderCallout1":
            case "accentBorderCallout2":
            case "accentBorderCallout3":
            case "accentCallout1":
            case "accentCallout2":
            case "accentCallout3":
            case "actionButtonBackPrevious":
            case "actionButtonBeginning":
            case "actionButtonBlank":
            case "actionButtonDocument":
            case "actionButtonEnd":
            case "actionButtonForwardNext":
            case "actionButtonHelp":
            case "actionButtonHome":
            case "actionButtonInformation":
            case "actionButtonMovie":
            case "actionButtonReturn":
            case "actionButtonSound":
            case "arc":
            case "bevel":
            case "blockArc":
            case "borderCallout1":
            case "borderCallout2":
            case "borderCallout3":
            case "bracePair":
            case "bracketPair":
            case "callout1":
            case "callout2":
            case "callout3":
            case "can":
            case "chartPlus":
            case "chartStar":
            case "chartX":
            case "chevron":
            case "chord":
            case "cloud":
            case "cloudCallout":
            case "corner":
            case "cornerTabs":
            case "cube":
            case "decagon":
            case "diagStripe":
            case "diamond":
            case "dodecagon":
            case "donut":
            case "doubleWave":
            case "downArrowCallout":
            case "ellipseRibbon":
            case "ellipseRibbon2":
            case "flowChartAlternateProcess":
            case "flowChartCollate":
            case "flowChartConnector":
            case "flowChartDecision":
            case "flowChartDelay":
            case "flowChartDisplay":
            case "flowChartDocument":
            case "flowChartExtract":
            case "flowChartInputOutput":
            case "flowChartInternalStorage":
            case "flowChartMagneticDisk":
            case "flowChartMagneticDrum":
            case "flowChartMagneticTape":
            case "flowChartManualInput":
            case "flowChartManualOperation":
            case "flowChartMerge":
            case "flowChartMultidocument":
            case "flowChartOfflineStorage":
            case "flowChartOffpageConnector":
            case "flowChartOnlineStorage":
            case "flowChartOr":
            case "flowChartPredefinedProcess":
            case "flowChartPreparation":
            case "flowChartProcess":
            case "flowChartPunchedCard":
            case "flowChartPunchedTape":
            case "flowChartSort":
            case "flowChartSummingJunction":
            case "flowChartTerminator":
            case "folderCorner":
            case "frame":
            case "funnel":
            case "gear6":
            case "gear9":
            case "halfFrame":
            case "heart":
            case "heptagon":
            case "hexagon":
            case "homePlate":
            case "horizontalScroll":
            case "irregularSeal1":
            case "irregularSeal2":
            case "leftArrow":
            case "leftArrowCallout":
            case "leftBrace":
            case "leftBracket":
            case "leftRightArrowCallout":
            case "leftRightRibbon":
            case "irregularSeal1":
            case "lightningBolt":
            case "lineInv":
            case "mathDivide":
            case "mathEqual":
            case "mathMinus":
            case "mathMultiply":
            case "mathNotEqual":
            case "mathPlus":
            case "moon":
            case "nonIsoscelesTrapezoid":
            case "noSmoking":
            case "octagon":
            case "parallelogram":
            case "pentagon":
            case "pie":
            case "pieWedge":
            case "plaque":
            case "plaqueTabs":
            case "plus":
            case "quadArrowCallout":
            case "rect":
            case "ribbon":
            case "ribbon2":
            case "rightArrowCallout":
            case "rightBrace":
            case "rightBracket":
            case "round1Rect":
            case "round2DiagRect":
            case "round2SameRect":
            case "rtTriangle":
            case "smileyFace":
            case "snip1Rect":
            case "snip2DiagRect":
            case "snip2SameRect":
            case "snipRoundRect":
            case "squareTabs":
            case "star10":
            case "star12":
            case "star16":
            case "star24":
            case "star32":
            case "star4":
            case "star5":
            case "star6":
            case "star7":
            case "star8":
            case "sun":
            case "teardrop":
            case "trapezoid":
            case "upArrowCallout":
            case "upDownArrowCallout":
            case "verticalScroll":
            case "wave":
            case "wedgeEllipseCallout":
            case "wedgeRectCallout":
            case "wedgeRoundRectCallout":
            case "rect":
                result += "<rect x='0' y='0' width='" + w + "' height='" + h + "' fill='" + fillColor + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                break;
            case "ellipse":
                result += "<ellipse cx='" + (w / 2) + "' cy='" + (h / 2) + "' rx='" + (w / 2) + "' ry='" + (h / 2) + "' fill='" + fillColor + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                break;
            case "roundRect":
                result += "<rect x='0' y='0' width='" + w + "' height='" + h + "' rx='7' ry='7' fill='" + fillColor + 
                            "' stroke='" + border.color + "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' />";
                break;
            case "bentConnector2":    // 直角 (path)
                var d = "";
                if (isFlipV) {
                    d = "M 0 " + w + " L " + h + " " + w + " L " + h + " 0";
                } else {
                    d = "M " + w + " 0 L " + w + " " + h + " L 0 " + h;
                }
                result += "<path d='" + d + "' stroke='" + border.color + 
                                "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' fill='none' ";
                if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                    result += "marker-start='url(#markerTriangle)' ";
                }
                if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                    result += "marker-end='url(#markerTriangle)' ";
                }
                result += "/>";
                break;
            case "line":
            case "straightConnector1":
            case "bentConnector3":
            case "bentConnector4":
            case "bentConnector5":
            case "curvedConnector2":
            case "curvedConnector3":
            case "curvedConnector4":
            case "curvedConnector5":
                if (isFlipV) {
                    result += "<line x1='" + w + "' y1='0' x2='0' y2='" + h + "' stroke='" + border.color + 
                                "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' ";
                } else {
                    result += "<line x1='0' y1='0' x2='" + w + "' y2='" + h + "' stroke='" + border.color + 
                                "' stroke-width='" + border.width + "' stroke-dasharray='" + border.strokeDasharray + "' ";
                }
                if (headEndNodeAttrs !== undefined && (headEndNodeAttrs["type"] === "triangle" || headEndNodeAttrs["type"] === "arrow")) {
                    result += "marker-start='url(#markerTriangle)' ";
                }
                if (tailEndNodeAttrs !== undefined && (tailEndNodeAttrs["type"] === "triangle" || tailEndNodeAttrs["type"] === "arrow")) {
                    result += "marker-end='url(#markerTriangle)' ";
                }
                result += "/>";
                break;
            case "rightArrow":
                result += "<defs><marker id=\"markerTriangle\" viewBox=\"0 0 10 10\" refX=\"1\" refY=\"5\" markerWidth=\"2.5\" markerHeight=\"2.5\" orient=\"auto-start-reverse\" markerUnits=\"strokeWidth\"><path d=\"M 0 0 L 10 5 L 0 10 z\" /></marker></defs>";
                result += "<line x1='0' y1='" + (h/2) + "' x2='" + (w-15) + "' y2='" + (h/2) + "' stroke='" + border.color + 
                                "' stroke-width='" + (h/2) + "' stroke-dasharray='" + border.strokeDasharray + "' ";
                result += "marker-end='url(#markerTriangle)' />";
                break;
            case "downArrow":
                result += "<defs><marker id=\"markerTriangle\" viewBox=\"0 0 10 10\" refX=\"1\" refY=\"5\" markerWidth=\"2.5\" markerHeight=\"2.5\" orient=\"auto-start-reverse\" markerUnits=\"strokeWidth\"><path d=\"M 0 0 L 10 5 L 0 10 z\" /></marker></defs>";
                result += "<line x1='" + (w/2) + "' y1='0' x2='" + (w/2) + "' y2='" + (h-15) + "' stroke='" + border.color + 
                                "' stroke-width='" + (w/2) + "' stroke-dasharray='" + border.strokeDasharray + "' ";
                result += "marker-end='url(#markerTriangle)' />";
                break;
            case "bentArrow":
            case "bentUpArrow":
            case "stripedRightArrow":
            case "quadArrow":
            case "circularArrow":
            case "swooshArrow":
            case "leftRightArrow":
            case "leftRightUpArrow":
            case "leftUpArrow":
            case "leftCircularArrow":
            case "notchedRightArrow":
            case "curvedDownArrow":
            case "curvedLeftArrow":
            case "curvedRightArrow":
            case "curvedUpArrow":
            case "upDownArrow":
            case "upArrow":
            case "uturnArrow":
            case "leftRightCircularArrow":
                break;
            case "triangle":
                break;
            case undefined:
            default:
                console.warn("Undefine shape type.");
        }
        
        result += "</svg>";
        
        result += "<div class='block content " + getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) +
                "' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
                "' style='" + 
                    getPosition(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) + 
                    getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) + 
                    " z-index: " + order + ";" +
                "'>";
        
        // TextBody
        if (node["p:txBody"] !== undefined) {
            result += genTextBody(node["p:txBody"], slideLayoutSpNode, slideMasterSpNode, type, warpObj);
        }
        result += "</div>";
        
    } else {
        
        result += "<div class='block content " + getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) +
                "' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
                "' style='" + 
                    getPosition(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) + 
                    getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) + 
                    getBorder(node, false) +
                    getShapeFill(node, false) +
                    " z-index: " + order + ";" +
                "'>";
        
        // TextBody
        if (node["p:txBody"] !== undefined) {
            result += genTextBody(node["p:txBody"], slideLayoutSpNode, slideMasterSpNode, type, warpObj);
        }
        result += "</div>";
        
    }
    
    return result;
}

function processPicNode(node, warpObj) {
    
    //debug( JSON.stringify( node ) );
    
    var order = node["attrs"]["order"];
    
    var rid = node["p:blipFill"]["a:blip"]["attrs"]["r:embed"];
    var imgName = warpObj["slideResObj"][rid]["target"];
    var imgFileExt = extractFileExtension(imgName).toLowerCase();
    var zip = warpObj["zip"];
    var imgArrayBuffer = zip.file(imgName).asArrayBuffer();
    var mimeType = "";
    var xfrmNode = node["p:spPr"]["a:xfrm"];
    switch (imgFileExt) {
        case "jpg":
        case "jpeg":
            mimeType = "image/jpeg";
            break;
        case "png":
            mimeType = "image/png";
            break;
        case "gif":
            mimeType = "image/gif";
            break;
        case "emf": // Not native support
            mimeType = "image/x-emf";
            break;
        case "wmf": // Not native support
            mimeType = "image/x-wmf";
            break;
        default:
            mimeType = "image/*";
    }
    return "<div class='block content' style='" + getPosition(xfrmNode, undefined, undefined) + getSize(xfrmNode, undefined, undefined) +
            " z-index: " + order + ";" +
            "'><img src=\"data:" + mimeType + ";base64," + base64ArrayBuffer(imgArrayBuffer) + "\" style='width: 100%; height: 100%'/></div>";
}

function processGraphicFrameNode(node, warpObj) {
    
    var result = "";
    var graphicTypeUri = getTextByPathList(node, ["a:graphic", "a:graphicData", "attrs", "uri"]);
    
    switch (graphicTypeUri) {
        case "http://schemas.openxmlformats.org/drawingml/2006/table":
            result = genTable(node, warpObj);
            break;
        case "http://schemas.openxmlformats.org/drawingml/2006/chart":
            result = genChart(node, warpObj);
            break;
        case "http://schemas.openxmlformats.org/drawingml/2006/diagram":
            result = genDiagram(node, warpObj);
            break;
        default:
    }
    
    return result;
}

function processSpPrNode(node, warpObj) {
    
    /*
     * 2241 <xsd:complexType name="CT_ShapeProperties">
     * 2242   <xsd:sequence>
     * 2243     <xsd:element name="xfrm" type="CT_Transform2D"  minOccurs="0" maxOccurs="1"/>
     * 2244     <xsd:group   ref="EG_Geometry"                  minOccurs="0" maxOccurs="1"/>
     * 2245     <xsd:group   ref="EG_FillProperties"            minOccurs="0" maxOccurs="1"/>
     * 2246     <xsd:element name="ln" type="CT_LineProperties" minOccurs="0" maxOccurs="1"/>
     * 2247     <xsd:group   ref="EG_EffectProperties"          minOccurs="0" maxOccurs="1"/>
     * 2248     <xsd:element name="scene3d" type="CT_Scene3D"   minOccurs="0" maxOccurs="1"/>
     * 2249     <xsd:element name="sp3d" type="CT_Shape3D"      minOccurs="0" maxOccurs="1"/>
     * 2250     <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
     * 2251   </xsd:sequence>
     * 2252   <xsd:attribute name="bwMode" type="ST_BlackWhiteMode" use="optional"/>
     * 2253 </xsd:complexType>
     */
    
    // TODO:
}

function genTextBody(textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj) {
    
    var text = "";
    var slideMasterTextStyles = warpObj["slideMasterTextStyles"];
    
    if (textBodyNode === undefined) {
        return text;
    }

    if (textBodyNode["a:p"].constructor === Array) {
        // multi p
        for (var i=0; i<textBodyNode["a:p"].length; i++) {
            var pNode = textBodyNode["a:p"][i];
            var rNode = pNode["a:r"];
            text += "<div class='" + getHorizontalAlign(pNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles) + "'>";
            text += genBuChar(pNode);
            if (rNode === undefined) {
                // without r
                text += genSpanElement(pNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj);
            } else if (rNode.constructor === Array) {
                // with multi r
                for (var j=0; j<rNode.length; j++) {
                    text += genSpanElement(rNode[j], slideLayoutSpNode, slideMasterSpNode, type, warpObj);
                }
            } else {
                // with one r
                text += genSpanElement(rNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj);
            }
            text += "</div>";
        }
    } else {
        // one p
        var pNode = textBodyNode["a:p"];
        var rNode = pNode["a:r"];
        text += "<div class='" + getHorizontalAlign(pNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles) + "'>";
        text += genBuChar(pNode);
        if (rNode === undefined) {
            // without r
            text += genSpanElement(pNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj);
        } else if (rNode.constructor === Array) {
            // with multi r
            for (var j=0; j<rNode.length; j++) {
                text += genSpanElement(rNode[j], slideLayoutSpNode, slideMasterSpNode, type, warpObj);
            }
        } else {
            // with one r
            text += genSpanElement(rNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj);
        }
        text += "</div>";
    }
    
    return text;
}

function genBuChar(node) {

    var pPrNode = node["a:pPr"];
    
    debug(JSON.stringify(pPrNode))
    
    var lvl = parseInt( getTextByPathList(pPrNode, ["attrs", "lvl"]) );
    if (isNaN(lvl)) {
        lvl = 0;
    }
    
    var buChar = getTextByPathList(pPrNode, ["a:buChar", "attrs", "char"]);
    if (buChar !== undefined) {
        var buFontAttrs = getTextByPathList(pPrNode, ["a:buFont", "attrs"]);
        if (buFontAttrs !== undefined) {
            var marginLeft = parseInt( getTextByPathList(pPrNode, ["attrs", "marL"]) ) * 96 / 914400;
            var marginRight = parseInt(buFontAttrs["pitchFamily"]);
            if (isNaN(marginLeft)) {
                marginLeft = 328600 * 96 / 914400;
            }
            if (isNaN(marginRight)) {
                marginRight = 0;
            }
            var typeface = buFontAttrs["typeface"];
            
            return "<span style='font-family: " + typeface + 
                    "; margin-left: " + marginLeft * lvl + "px" +
                    "; margin-right: " + marginRight + "px" +
                    "; font-size: 20pt" +
                    "'>" + buChar + "</span>";
        } else {
            marginLeft = 328600 * 96 / 914400 * lvl;
            return "<span style='margin-left: " + marginLeft + "px;'>" + buChar + "</span>";
        }
    } else {
        //buChar = '•';
        return "<span style='margin-left: " + 328600 * 96 / 914400 * lvl + "px" +
                    "; margin-right: " + 0 + "px;'></span>";
    }
    
    return "";
}

function genSpanElement(node, slideLayoutSpNode, slideMasterSpNode, type, warpObj) {
    
    var slideMasterTextStyles = warpObj["slideMasterTextStyles"];
    
    var text = node["a:t"];
    if (typeof text !== 'string') {
        text = getTextByPathList(node, ["a:fld", "a:t"]);
        if (typeof text !== 'string') {
            text = "&nbsp;";
            //debug("XXX: " + JSON.stringify(node));
        }
    }
    
    var styleText = 
        "color:" + getFontColor(node, type, slideMasterTextStyles) + 
        ";font-size:" + getFontSize(node, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles) + 
        ";font-family:" + getFontType(node, type, slideMasterTextStyles) + 
        ";font-weight:" + getFontBold(node, type, slideMasterTextStyles) + 
        ";font-style:" + getFontItalic(node, type, slideMasterTextStyles) + 
        ";text-decoration:" + getFontDecoration(node, type, slideMasterTextStyles) +
        ";vertical-align:" + getTextVerticalAlign(node, type, slideMasterTextStyles) + 
        ";";
    
    var cssName = "";
    
    if (styleText in styleTable) {
        cssName = styleTable[styleText]["name"];
    } else {
        cssName = "_css_" + (Object.keys(styleTable).length + 1);
        styleTable[styleText] = {
            "name": cssName,
            "text": styleText
        };
    }
    
    var linkID = getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "r:id"]);
    if (linkID !== undefined) {
        var linkURL = warpObj["slideResObj"][linkID]["target"];
        return "<span class='text-block " + cssName + "'><a href='" + linkURL + "' target='_blank'>" + text.replace(/\s/i, "&nbsp;") + "</a></span>";
    } else {
        return "<span class='text-block " + cssName + "'>" + text.replace(/\s/i, "&nbsp;") + "</span>";
    }
    
}

function genGlobalCSS() {
    var cssText = "";
    for (var key in styleTable) {
        cssText += "section ." + styleTable[key]["name"] + "{" + styleTable[key]["text"] + "}\n";
    }
    return cssText;
}

function genTable(node, warpObj) {
    
    var order = node["attrs"]["order"];
    var tableNode = getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl"]);
    var xfrmNode = getTextByPathList(node, ["p:xfrm"]);
    var tableHtml = "<table style='" + getPosition(xfrmNode, undefined, undefined) + getSize(xfrmNode, undefined, undefined) + " z-index: " + order + ";'>";
    
    var trNodes = tableNode["a:tr"];
    if (trNodes.constructor === Array) {
        for (var i=0; i<trNodes.length; i++) {
            tableHtml += "<tr>";
            var tcNodes = trNodes[i]["a:tc"];
            
            if (tcNodes.constructor === Array) {
                for (var j=0; j<tcNodes.length; j++) {
                    var text = genTextBody(tcNodes[j]["a:txBody"], undefined, undefined, undefined, warpObj);        
                    var rowSpan = getTextByPathList(tcNodes[j], ["attrs", "rowSpan"]);
                    var colSpan = getTextByPathList(tcNodes[j], ["attrs", "gridSpan"]);
                    var vMerge = getTextByPathList(tcNodes[j], ["attrs", "vMerge"]);
                    var hMerge = getTextByPathList(tcNodes[j], ["attrs", "hMerge"]);
                    if (rowSpan !== undefined) {
                        tableHtml += "<td rowspan='" + parseInt(rowSpan) + "'>" + text + "</td>";
                    } else if (colSpan !== undefined) {
                        tableHtml += "<td colspan='" + parseInt(colSpan) + "'>" + text + "</td>";
                    } else if (vMerge === undefined && hMerge === undefined) {
                        tableHtml += "<td>" + text + "</td>";
                    }
                }
            } else {
                var text = genTextBody(tcNodes["a:txBody"]);
                tableHtml += "<td>" + text + "</td>";
            }
            tableHtml += "</tr>";
        }
    } else {
        tableHtml += "<tr>";
        var tcNodes = trNodes["a:tc"];
        if (tcNodes.constructor === Array) {
            for (var j=0; j<tcNodes.length; j++) {
                var text = genTextBody(tcNodes[j]["a:txBody"]);
                tableHtml += "<td>" + text + "</td>";
            }
        } else {
            var text = genTextBody(tcNodes["a:txBody"]);
            tableHtml += "<td>" + text + "</td>";
        }
        tableHtml += "</tr>";
    }
    
    return tableHtml;
}

function genChart(node, warpObj) {
    
    var order = node["attrs"]["order"];
    var xfrmNode = getTextByPathList(node, ["p:xfrm"]);
    var result = "<div id='chart" + chartID + "' class='block content' style='" + 
                    getPosition(xfrmNode, undefined, undefined) + getSize(xfrmNode, undefined, undefined) + 
                    " z-index: " + order + ";'></div>";
    
    var rid = node["a:graphic"]["a:graphicData"]["c:chart"]["attrs"]["r:id"];
    var refName = warpObj["slideResObj"][rid]["target"];
    var content = readXmlFile(warpObj["zip"], refName);
    var plotArea = getTextByPathList(content, ["c:chartSpace", "c:chart", "c:plotArea"]);
    
    var chartData = null;
    for (var key in plotArea) {
        switch (key) {
            case "c:lineChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartID": "chart" + chartID,
                        "chartType": "lineChart",
                        "chartData": extractChartData(plotArea[key]["c:ser"])
                    }
                };
                break;
            case "c:barChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartID": "chart" + chartID,
                        "chartType": "barChart",
                        "chartData": extractChartData(plotArea[key]["c:ser"])
                    }
                };
                break;
            case "c:pieChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartID": "chart" + chartID,
                        "chartType": "pieChart",
                        "chartData": extractChartData(plotArea[key]["c:ser"])
                    }
                };
                break;
            case "c:pie3DChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartID": "chart" + chartID,
                        "chartType": "pie3DChart",
                        "chartData": extractChartData(plotArea[key]["c:ser"])
                    }
                };
                break;
            case "c:areaChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartID": "chart" + chartID,
                        "chartType": "areaChart",
                        "chartData": extractChartData(plotArea[key]["c:ser"])
                    }
                };
                break;
            case "c:scatterChart":
                chartData = {
                    "type": "createChart",
                    "data": {
                        "chartID": "chart" + chartID,
                        "chartType": "scatterChart",
                        "chartData": extractChartData(plotArea[key]["c:ser"])
                    }
                };
                break;
            case "c:catAx":
                break;
            case "c:valAx":
                break;
            default:
        }
    }
    
    if (chartData !== null) {
        MsgQueue.push(chartData);
    }
    
    chartID++;
    return result;
}

function genDiagram(node, warpObj) {
    var order = node["attrs"]["order"];
    var xfrmNode = getTextByPathList(node, ["p:xfrm"]);
    return "<div class='block content' style='border: 1px dotted;" + 
                getPosition(xfrmNode, undefined, undefined) + getSize(xfrmNode, undefined, undefined) + 
            "'>TODO: diagram</div>";
}

function getPosition(slideSpNode, slideLayoutSpNode, slideMasterSpNode) {
    
    //debug(JSON.stringify(slideLayoutSpNode));
    //debug(JSON.stringify(slideMasterSpNode));
    
    var off = undefined;
    var x = -1, y = -1;
    
    if (slideSpNode !== undefined) {
        off = slideSpNode["a:off"]["attrs"];
    } else if (slideLayoutSpNode !== undefined) {
        off = slideLayoutSpNode["a:off"]["attrs"];
    } else if (slideMasterSpNode !== undefined) {
        off = slideMasterSpNode["a:off"]["attrs"];
    }
    
    if (off === undefined) {
        return "";
    } else {
        x = parseInt(off["x"]) * 96 / 914400;
        y = parseInt(off["y"]) * 96 / 914400;
        return (isNaN(x) || isNaN(y)) ? "" : "top:" + y + "px; left:" + x + "px;";
    }
    
}

function getSize(slideSpNode, slideLayoutSpNode, slideMasterSpNode) {
    
    //debug(JSON.stringify(slideLayoutSpNode));
    //debug(JSON.stringify(slideMasterSpNode));
    
    var ext = undefined;
    var w = -1, h = -1;
    
    if (slideSpNode !== undefined) {
        ext = slideSpNode["a:ext"]["attrs"];
    } else if (slideLayoutSpNode !== undefined) {
        ext = slideLayoutSpNode["a:ext"]["attrs"];
    } else if (slideMasterSpNode !== undefined) {
        ext = slideMasterSpNode["a:ext"]["attrs"];
    }
    
    if (ext === undefined) {
        return "";
    } else {
        w = parseInt(ext["cx"]) * 96 / 914400;
        h = parseInt(ext["cy"]) * 96 / 914400;
        return (isNaN(w) || isNaN(h)) ? "" : "width:" + w + "px; height:" + h + "px;";
    }    
    
}

function getHorizontalAlign(node, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles) {
    //debug(node);
    var algn = getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
    if (algn === undefined) {
        algn = getTextByPathList(slideLayoutSpNode, ["p:txBody", "a:p", "a:pPr", "attrs", "algn"]);
        if (algn === undefined) {
            algn = getTextByPathList(slideMasterSpNode, ["p:txBody", "a:p", "a:pPr", "attrs", "algn"]);
            if (algn === undefined) {
                switch (type) {
                    case "title":
                    case "subTitle":
                    case "ctrTitle":
                        algn = getTextByPathList(slideMasterTextStyles, ["p:titleStyle", "a:lvl1pPr", "attrs", "alng"]);
                        break;
                    default:
                        algn = getTextByPathList(slideMasterTextStyles, ["p:otherStyle", "a:lvl1pPr", "attrs", "alng"]);
                }
            }
        }
    }
    // TODO:
    if (algn === undefined) {
        if (type == "title" || type == "subTitle" || type == "ctrTitle") {
            return "h-mid";
        } else if (type == "sldNum") {
            return "h-right";
        }
    }
    return algn === "ctr" ? "h-mid" : algn === "r" ? "h-right" : "h-left";
}

function getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles) {
    
    // 上中下對齊: X, <a:bodyPr anchor="ctr">, <a:bodyPr anchor="b">
    var anchor = getTextByPathList(node, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
    if (anchor === undefined) {
        anchor = getTextByPathList(slideLayoutSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
        if (anchor === undefined) {
            anchor = getTextByPathList(slideMasterSpNode, ["p:txBody", "a:bodyPr", "attrs", "anchor"]);
        }
    }
    
    return anchor === "ctr" ? "v-mid" : anchor === "b" ?  "v-down" : "v-up";
}

function getFontType(node, type, slideMasterTextStyles) {
    var typeface = getTextByPathList(node, ["a:rPr", "a:latin", "attrs", "typeface"]);
    
    if (typeface === undefined) {
        var fontSchemeNode = getTextByPathList(themeContent, ["a:theme", "a:themeElements", "a:fontScheme"]);
        if (type == "title" || type == "subTitle" || type == "ctrTitle") {
            typeface = getTextByPathList(fontSchemeNode, ["a:majorFont", "a:latin", "attrs", "typeface"]);
        } else if (type == "body") {
            typeface = getTextByPathList(fontSchemeNode, ["a:minorFont", "a:latin", "attrs", "typeface"]);
        } else {
            typeface = getTextByPathList(fontSchemeNode, ["a:minorFont", "a:latin", "attrs", "typeface"]);
        }
    }
    
    return (typeface === undefined) ? "inherit" : typeface;
}

function getFontColor(node, type, slideMasterTextStyles) {
    var color = getTextByPathStr(node, "a:rPr a:solidFill a:srgbClr attrs val");
    return (color === undefined) ? "#000" : "#" + color;
}

function getFontSize(node, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles) {
    var fontSize = undefined;
    if (node["a:rPr"] !== undefined) {
        fontSize = parseInt(node["a:rPr"]["attrs"]["sz"]) / 100;
    }
    
    if ((isNaN(fontSize) || fontSize === undefined)) {
        var sz = getTextByPathList(slideLayoutSpNode, ["p:txBody", "a:lstStyle", "a:lvl1pPr", "a:defRPr", "attrs", "sz"]);
        fontSize = parseInt(sz) / 100;
    }
    
    if (isNaN(fontSize) || fontSize === undefined) {
        if (type == "title" || type == "subTitle" || type == "ctrTitle") {
            var sz = getTextByPathList(slideMasterTextStyles, ["p:titleStyle", "a:lvl1pPr", "a:defRPr", "attrs", "sz"]);
        } else if (type == "body") {
            var sz = getTextByPathList(slideMasterTextStyles, ["p:bodyStyle", "a:lvl1pPr", "a:defRPr", "attrs", "sz"]);
        } else if (type == "dt" || type == "sldNum") {
            var sz = "1200";
        } else if (type === undefined) {
            var sz = getTextByPathList(slideMasterTextStyles, ["p:otherStyle", "a:lvl1pPr", "a:defRPr", "attrs", "sz"]);
        }
        fontSize = parseInt(sz) / 100;
    }
    
    var baseline = getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
    if (baseline !== undefined && !isNaN(fontSize)) {
        fontSize -= 10;
    }
    
    return isNaN(fontSize) ? "inherit" : (fontSize + "pt");
}

function getFontBold(node, type, slideMasterTextStyles) {
    return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["b"] === "1") ? "bold" : "initial";
}

function getFontItalic(node, type, slideMasterTextStyles) {
    return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["i"] === "1") ? "italic" : "normal";
}

function getFontDecoration(node, type, slideMasterTextStyles) {
    return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["u"] === "sng") ? "underline" : "initial";
}

function getTextVerticalAlign(node, type, slideMasterTextStyles) {
    var baseline = getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
    return baseline === undefined ? "baseline" : (parseInt(baseline) / 1000) + "%";
}

function getBorder(node, isSvgMode) {
    
    //debug(JSON.stringify(node));
    
    var cssText = "border: ";
    
    // 1. presentationML
    var lineNode = node["p:spPr"]["a:ln"];
    
    // Border width: 1pt = 12700, default = 0.75pt
    var borderWidth = parseInt(getTextByPathList(lineNode, ["attrs", "w"])) / 12700;
    if (isNaN(borderWidth) || borderWidth < 1) {
        cssText += "1pt ";
    } else {
        cssText += borderWidth + "pt ";
    }
    
    // Border color
    var borderColor = getTextByPathList(lineNode, ["a:solidFill", "a:srgbClr", "attrs", "val"]);
    if (borderColor === undefined) {
        var schemeClrNode = getTextByPathList(lineNode, ["a:solidFill", "a:schemeClr"]);
        var schemeClr = "a:" + getTextByPathList(schemeClrNode, ["attrs", "val"]);    
        var borderColor = getSchemeColorFromTheme(schemeClr);
    }
    
    // 2. drawingML namespace
    if (borderColor === undefined) {
        var schemeClrNode = getTextByPathList(node, ["p:style", "a:lnRef", "a:schemeClr"]);
        var schemeClr = "a:" + getTextByPathList(schemeClrNode, ["attrs", "val"]);    
        var borderColor = getSchemeColorFromTheme(schemeClr);
        
        if (borderColor !== undefined) {
            var shade = getTextByPathList(schemeClrNode, ["a:shade", "attrs", "val"]);
            if (shade !== undefined) {
                shade = parseInt(shade) / 100000;
                var color = new colz.Color("#" + borderColor);
                color.setLum(color.hsl.l * shade);
                borderColor = color.hex.replace("#", "");
            }
        }
        
    }
    
    if (borderColor === undefined) {
        if (isSvgMode) {
            borderColor = "none";
        } else {
            borderColor = "#000";
        }
    } else {
        borderColor = "#" + borderColor;
        
    }
    cssText += " " + borderColor + " ";
    
    // Border type
    var borderType = getTextByPathList(lineNode, ["a:prstDash", "attrs", "val"]);
    var strokeDasharray = "0";
    switch (borderType) {
        case "solid":
            cssText += "solid";
            strokeDasharray = "0";
            break;
        case "dash":
            cssText += "dashed";
            strokeDasharray = "5";
            break;
        case "dashDot":
            cssText += "dashed";
            strokeDasharray = "5, 5, 1, 5";
            break;
        case "dot":
            cssText += "dotted";
            strokeDasharray = "1, 5";
            break;
        case "lgDash":
            cssText += "dashed";
            strokeDasharray = "10, 5";
            break;
        case "lgDashDotDot":
            cssText += "dashed";
            strokeDasharray = "10, 5, 1, 5, 1, 5";
            break;
        case "sysDash":
            cssText += "dashed";
            strokeDasharray = "5, 2";
            break;
        case "sysDashDot":
            cssText += "dashed";
            strokeDasharray = "5, 2, 1, 5";
            break;
        case "sysDashDotDot":
            cssText += "dashed";
            strokeDasharray = "5, 2, 1, 5, 1, 5";
            break;
        case "sysDot":
            cssText += "dotted";
            strokeDasharray = "2, 5";
            break;
        case undefined:
            //console.log(borderType);
        default:
            //console.warn(borderType);
            //cssText += "#000 solid";
    }
    
    if (isSvgMode) {
        return {"color": borderColor, "width": borderWidth, "type": borderType, "strokeDasharray": strokeDasharray};
    } else {
        return cssText + ";";
    }
}

function getSlideBackgroundFill(slideContent, slideLayoutContent, slideMasterContent) {
    var bgColor = getSolidFill( getTextByPathList(slideContent, ["p:sld", "p:cSld", "p:bg", "p:bgPr", "a:solidFill"]) );
    if (bgColor === undefined) {
        bgColor = getSolidFill( getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:cSld", "p:bg", "p:bgPr", "a:solidFill"]) );
        if (bgColor === undefined) {
            bgColor = getSolidFill( getTextByPathList(slideMasterContent, ["p:sldMaster", "p:cSld", "p:bg", "p:bgPr", "a:solidFill"]) );
            if (bgColor === undefined) {
                bgColor = "FFF";
            }
        }
    }
    return bgColor;
}

function getShapeFill(node, isSvgMode) {
    
    // 1. presentationML
    // p:spPr [a:noFill, solidFill, gradFill, blipFill, pattFill, grpFill]
    // From slide
    if (getTextByPathList(node, ["p:spPr", "a:noFill"]) !== undefined) {
        return isSvgMode ? "none" : "background-color: initial;";
    }
    
    var fillColor = undefined;
    if (fillColor === undefined) {
        fillColor = getTextByPathList(node, ["p:spPr", "a:solidFill", "a:srgbClr", "attrs", "val"]);
    }
    
    // From theme
    if (fillColor === undefined) {
        var schemeClr = "a:" + getTextByPathList(node, ["p:spPr", "a:solidFill", "a:schemeClr", "attrs", "val"]);
        fillColor = getSchemeColorFromTheme(schemeClr);
    }
    
    // 2. drawingML namespace
    if (fillColor === undefined) {
        var schemeClr = "a:" + getTextByPathList(node, ["p:style", "a:fillRef", "a:schemeClr", "attrs", "val"]);
        fillColor = getSchemeColorFromTheme(schemeClr);
    }
    
    if (fillColor !== undefined) {
        
        fillColor = "#" + fillColor;
        
        // Apply shade or tint
        // TODO: 較淺, 較深 80%
        var lumMod = parseInt(getTextByPathList(node, ["p:spPr", "a:solidFill", "a:schemeClr", "a:lumMod", "attrs", "val"])) / 100000;
        var lumOff = parseInt(getTextByPathList(node, ["p:spPr", "a:solidFill", "a:schemeClr", "a:lumOff", "attrs", "val"])) / 100000;
        if (isNaN(lumMod)) {
            lumMod = 1.0;
        }
        if (isNaN(lumOff)) {
            lumOff = 0;
        }
        //console.log([lumMod, lumOff]);
        fillColor = applyLumModify(fillColor, lumMod, lumOff);
        
        if (isSvgMode) {
            return fillColor;
        } else {
            return "background-color: " + fillColor + ";";
        }
    } else {
        if (isSvgMode) {
            return "none";
        } else {
            return "background-color: " + fillColor + ";";
        }
        
    }
    
}

function getSolidFill(solidFill) {
    
    if (solidFill === undefined) {
        return undefined;
    }
    
    var color = "FFF";
    
    if (solidFill["a:srgbClr"] !== undefined) {
        color = getTextByPathList(solidFill["a:srgbClr"], ["attrs", "val"]);
    } else if (solidFill["a:schemeClr"] !== undefined) {
        var schemeClr = "a:" + getTextByPathList(solidFill["a:schemeClr"], ["attrs", "val"]);
        color = getSchemeColorFromTheme(schemeClr);
    }
    
    return color;
}

function getSchemeColorFromTheme(schemeClr) {
    // TODO: <p:clrMap ...> in slide master
    // e.g. tx2="dk2" bg2="lt2" tx1="dk1" bg1="lt1"
    switch (schemeClr) {
        case "a:tx1": schemeClr = "a:dk1"; break;
        case "a:tx2": schemeClr = "a:dk2"; break;
        case "a:bg1": schemeClr = "a:lt1"; break;
        case "a:bg2": schemeClr = "a:lt2"; break;
    }
    var refNode = getTextByPathList(themeContent, ["a:theme", "a:themeElements", "a:clrScheme", schemeClr]);
    var color = getTextByPathList(refNode, ["a:srgbClr", "attrs", "val"]);
    if (color === undefined) {
        color = getTextByPathList(refNode, ["a:sysClr", "attrs", "lastClr"]);
    }
    return color;
}

function extractChartData(serNode) {
    
    var dataMat = new Array();
    
    if (serNode === undefined) {
        return dataMat;
    }
    
    if (serNode["c:xVal"] !== undefined) {
        var dataRow = new Array();
        eachElement(serNode["c:xVal"]["c:numRef"]["c:numCache"]["c:pt"], function(innerNode, index) {
            dataRow.push(parseFloat(innerNode["c:v"]));
            return "";
        });
        dataMat.push(dataRow);
        dataRow = new Array();
        eachElement(serNode["c:yVal"]["c:numRef"]["c:numCache"]["c:pt"], function(innerNode, index) {
            dataRow.push(parseFloat(innerNode["c:v"]));
            return "";
        });
        dataMat.push(dataRow);
    } else {
        eachElement(serNode, function(innerNode, index) {
            var dataRow = new Array();
            var colName = getTextByPathList(innerNode, ["c:tx", "c:strRef", "c:strCache", "c:pt", "c:v"]) || index;

            // Category (string or number)
            var rowNames = {};
            if (getTextByPathList(innerNode, ["c:cat", "c:strRef", "c:strCache", "c:pt"]) !== undefined) {
                eachElement(innerNode["c:cat"]["c:strRef"]["c:strCache"]["c:pt"], function(innerNode, index) {
                    rowNames[innerNode["attrs"]["idx"]] = innerNode["c:v"];
                    return "";
                });
            } else if (getTextByPathList(innerNode, ["c:cat", "c:numRef", "c:numCache", "c:pt"]) !== undefined) {
                eachElement(innerNode["c:cat"]["c:numRef"]["c:numCache"]["c:pt"], function(innerNode, index) {
                    rowNames[innerNode["attrs"]["idx"]] = innerNode["c:v"];
                    return "";
                });
            }
            
            // Value
            if (getTextByPathList(innerNode, ["c:val", "c:numRef", "c:numCache", "c:pt"]) !== undefined) {
                eachElement(innerNode["c:val"]["c:numRef"]["c:numCache"]["c:pt"], function(innerNode, index) {
                    dataRow.push({x: innerNode["attrs"]["idx"], y: parseFloat(innerNode["c:v"])});
                    return "";
                });
            }
            
            dataMat.push({key: colName, values: dataRow, xlabels: rowNames});
            return "";
        });
    }
    
    return dataMat;
}

// ===== Node functions =====
/**
 * getTextByPathStr
 * @param {Object} node
 * @param {string} pathStr
 */
function getTextByPathStr(node, pathStr) {
    return getTextByPathList(node, pathStr.trim().split(/\s+/));
}

/**
 * getTextByPathList
 * @param {Object} node
 * @param {string Array} path
 */
function getTextByPathList(node, path) {

    if (path.constructor !== Array) {
        throw Error("Error of path type! path is not array.");
    }
    
    if (node === undefined) {
        return undefined;
    }
    
    var l = path.length;
    for (var i=0; i<l; i++) {
        node = node[path[i]];
        if (node === undefined) {
            return undefined;
        }
    }
    
    return node;
}

/**
 * eachElement
 * @param {Object} node
 * @param {function} doFunction
 */
function eachElement(node, doFunction) {
    if (node === undefined) {
        return;
    }
    var result = "";
    if (node.constructor === Array) {
        var l = node.length;
        for (var i=0; i<l; i++) {
            result += doFunction(node[i], i);
        }
    } else {
        result += doFunction(node, 0);
    }
    return result;
}

// ===== Color functions =====
/**
 * applyShade
 * @param {string} rgbStr
 * @param {number} shadeValue
 */
function applyShade(rgbStr, shadeValue) {
    var color = new colz.Color(rgbStr);
    color.setLum(color.hsl.l * shadeValue);
    return color.rgb.toString();
}

/**
 * applyTint
 * @param {string} rgbStr
 * @param {number} tintValue
 */
function applyTint(rgbStr, tintValue) {
    var color = new colz.Color(rgbStr);
    color.setLum(color.hsl.l * tintValue + (1 - tintValue));
    return color.rgb.toString();
}

/**
 * applyLumModify
 * @param {string} rgbStr
 * @param {number} factor
 * @param {number} offset
 */
function applyLumModify(rgbStr, factor, offset) {
    var color = new colz.Color(rgbStr);
    //color.setLum(color.hsl.l * factor);
    color.setLum(color.hsl.l * (1 + offset));
    return color.rgb.toString();
}

// ===== Debug functions =====
/**
 * debug
 * @param {Object} data
 */
function debug(data) {
    self.postMessage({"type": "DEBUG", "data": data});
}
