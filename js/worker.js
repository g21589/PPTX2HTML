"use strict";

importScripts(
	'./jszip.min.js',
	'./highlight.min.js',
	'./colz.class.min.js',
	'./highlight.min.js',
	'./tXml.js',
	'./functions.js'
);

var themeContent = null;

var titleFontSize = 42;
var bodyFontSize = 20;
var otherFontSize = 16;

onmessage = function(e) {
	
	var dateBefore = new Date();
	
	var zip = new JSZip(e.data);
	
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
	var slideMasterTables = indexNodes(slideMasterContent);
	//debug(slideMasterTables);
	
	
	// =====< Step 3 >=====
	var content = readXmlFile(zip, sldFileName);
	var nodes = content["p:sld"]["p:cSld"]["p:spTree"];
	var warpObj = {
		"zip": zip,
		"slideLayoutTables": slideLayoutTables,
		"slideMasterTables": slideMasterTables,
		"slideResObj": slideResObj
	};
	
	var result = "<li class='slide'>" + sldFileName + "<section style='width:" + slideSize.width + "px; height:" + slideSize.height + "px;'>"
	
	for (var nodeKey in nodes) {
		if (nodes[nodeKey].constructor === Array) {
			for (var i=0; i<nodes[nodeKey].length; i++) {
				result += processNodesInSlide(nodeKey, nodes[nodeKey][i], warpObj, 0);
			}
		} else {
			result += processNodesInSlide(nodeKey, nodes[nodeKey], warpObj, 0);
		}
	}
	
	return result + "</section></li>";
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

function processNodesInSlide(nodeKey, nodeValue, warpObj, depth) {
	
	var result = "";
	
	switch (nodeKey) {
		case "p:sp":	// Shape, Text
			result += processSpNode(nodeValue, warpObj);
			break;
		case "p:cxnSp":	// Shape, Text (with connection)
			result += processCxnSpNode(nodeValue, warpObj);
			break;
		case "p:pic":	// Picture
			result += processPicNode(nodeValue, warpObj);
			break;
		case "p:graphicFrame":	// Chart, Diagram, Table
			result += processGraphicFrameNode(nodeValue, warpObj);
			break;
		case "p:grpSp":	// 群組
			var order = nodeValue["attrs"]["order"];
			result += "<div class='block group' style='z-index: " + order + ";";			
			for (var nodeKey in nodeValue) {
				if (nodeValue[nodeKey].constructor === Array) {
					for (var i=0; i<nodeValue[nodeKey].length; i++) {
						result += processNodesInSlide(nodeKey, nodeValue[nodeKey][i], warpObj, depth + 1);
					}
				} else {
					result += processNodesInSlide(nodeKey, nodeValue[nodeKey], warpObj, depth + 1);
				}
			}
			result += "</div>";
			break;
		case "p:nvGrpSpPr":
			// id
			//$node.find("cNvPr").attr("id");
			break;
		case "p:grpSpPr":
			// size
			if (depth > 0) {
				var xfrmNode = nodeValue["a:xfrm"];
				var x = parseInt(xfrmNode["a:off"]["attrs"]["x"]) * 96 / 914400;
				var y = parseInt(xfrmNode["a:off"]["attrs"]["y"]) * 96 / 914400;
				var chx = parseInt(xfrmNode["a:chOff"]["attrs"]["x"]) * 96 / 914400;
				var chy = parseInt(xfrmNode["a:chOff"]["attrs"]["y"]) * 96 / 914400;
				var cx = parseInt(xfrmNode["a:ext"]["attrs"]["cx"]) * 96 / 914400;
				var cy = parseInt(xfrmNode["a:ext"]["attrs"]["cy"]) * 96 / 914400;
				var chcx = parseInt(xfrmNode["a:chExt"]["attrs"]["cx"]) * 96 / 914400;
				var chcy = parseInt(xfrmNode["a:chExt"]["attrs"]["cy"]) * 96 / 914400;
				result = " top: " + (y - chy) + "px; left: " + (x - chx) + "px; width: " + (cx - chcx) + "px; height: " + (cy - chcy) + "px;'>";
			}
			break;
		default:
	}
	
	return result;
	
}

function processSpNode(node, warpObj) {
	
	/*
	 *  958	<xsd:complexType name="CT_GvmlShape">
	 *  959   <xsd:sequence>
	 *  960     <xsd:element name="nvSpPr" type="CT_GvmlShapeNonVisual"     minOccurs="1" maxOccurs="1"/>
	 *  961     <xsd:element name="spPr"   type="CT_ShapeProperties"        minOccurs="1" maxOccurs="1"/>
	 *  962     <xsd:element name="txSp"   type="CT_GvmlTextShape"          minOccurs="0" maxOccurs="1"/>
	 *  963     <xsd:element name="style"  type="CT_ShapeStyle"             minOccurs="0" maxOccurs="1"/>
	 *  964     <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
	 *  965   </xsd:sequence>
	 *  966 </xsd:complexType>
	 */
	
	var id = node["p:nvSpPr"]["p:cNvSpPr"]["attrs"]["id"];
	var name = node["p:nvSpPr"]["p:cNvSpPr"]["attrs"]["name"];
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
	
	return genShape(node, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order);
}

function processCxnSpNode(node, warpObj) {

	var id = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["id"];
	var name = node["p:nvCxnSpPr"]["p:cNvPr"]["attrs"]["name"];
	//var idx = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["idx"];
	//var type = (node["p:nvCxnSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["type"];
	//<p:cNvCxnSpPr>(<p:cNvCxnSpPr>, <a:endCxn>)
	var order = node["attrs"]["order"];

	debug( {"id": id, "name": name, "order": order} );
	
	return genShape(node, undefined, undefined, id, name, undefined, undefined, order);
}

function genShape(node, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, order) {
	
	var xfrmList = ["p:spPr", "a:xfrm"];
	var slideXfrmNode = getTextByPathList(node, xfrmList);
	var slideLayoutXfrmNode = getTextByPathList(slideLayoutSpNode, xfrmList);
	var slideMasterXfrmNode = getTextByPathList(slideMasterSpNode, xfrmList);
	
	var result = "";
	var svgMode = node["p:style"] !== undefined;
	
	var isFlipV = false;
	if ( getTextByPathList(slideXfrmNode, ["attrs", "flipV"]) === "1" || getTextByPathList(slideXfrmNode, ["attrs", "flipH"]) === "1") {
		isFlipV = true;
	}
	
	// TODO: 圖形統一用SVG渲染
	// 優先權: p:spPr(a:solidFill, a:ln, ...), p:style(a:fillRef, a:lnRef, ...)
	if (svgMode) {
		
		var off = slideXfrmNode["a:off"]["attrs"];
		var x = parseInt(off["x"]) * 96 / 914400;
		var y = parseInt(off["y"]) * 96 / 914400;
		
		var ext = slideXfrmNode["a:ext"]["attrs"];
		var w = parseInt(ext["cx"]) * 96 / 914400;
		var h = parseInt(ext["cy"]) * 96 / 914400;
		
		result += "<svg class='drawing' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
				"' style='" + 
					getPosition(slideXfrmNode, undefined, undefined) + 
					getSize(slideXfrmNode, undefined, undefined) +
					" z-index: " + order + ";" +
				"'>";
		
		// Fill Color
		var fillColor = getFill(node, true);
		
		// Border Color		
		var border = getBorder(node, true);
		
		var headEndNodeAttrs = getTextByPathList(node, ["p:spPr", "a:ln", "a:headEnd", "attrs"]);
		var tailEndNodeAttrs = getTextByPathList(node, ["p:spPr", "a:ln", "a:tailEnd", "attrs"]);
		if ((headEndNodeAttrs !== undefined && headEndNodeAttrs["type"] === "triangle") || 
			(tailEndNodeAttrs !== undefined && tailEndNodeAttrs["type"] === "triangle")) {
				var triangleMarker = "<defs><marker id=\"markerTriangle\" viewBox=\"0 0 10 10\" refX=\"1\" refY=\"5\" markerWidth=\"5\" markerHeight=\"5\" orient=\"auto\"><path d=\"M 0 0 L 10 5 L 0 10 z\" /></marker></defs>";
				result += triangleMarker;
		}
		
		switch ( getTextByPathList(node, ["p:spPr", "a:prstGeom", "attrs", "prst"]) ) {
			case "rect":
				result += "<rect x='0' y='0' width='" + w + "' height='" + h + "' fill='" + fillColor + "' stroke='" + border.color + "' stroke-width='" + border.width + "' />";
				break;
			case "ellipse":
				result += "<ellipse cx='" + (w / 2) + "' cy='" + (h / 2) + "' rx='" + (w / 2) + "' ry='" + (h / 2) + "' fill='" + fillColor + "' stroke='" + border.color + "' stroke-width='" + border.width + "' />";
				break;
			case "roundRect":
				result += "<rect x='0' y='0' width='" + w + "' height='" + h + "' rx='15' ry='15' fill='" + fillColor + "' stroke='" + border.color + "' stroke-width='" + border.width + "' />";
				break;
			case "line":
			case "straightConnector1":
				if (isFlipV) {
					result += "<line x1='" + w + "' y1='0' x2='0' y2='" + h + "' stroke='" + border.color + "' stroke-width='" + border.width + "' ";
				} else {
					result += "<line x1='0' y1='0' x2='" + w + "' y2='" + h + "' stroke='" + border.color + "' stroke-width='" + border.width + "' ";
				}
				if (headEndNodeAttrs !== undefined && headEndNodeAttrs["type"] === "triangle") {
					result += "marker-start='url(#markerTriangle)' ";
				}
				if (tailEndNodeAttrs !== undefined && tailEndNodeAttrs["type"] === "triangle") {
					result += "marker-end='url(#markerTriangle)' ";
				}
				result += "/>";
				break;
			case "rightArrow":
			case "triangle":
			default:
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
			result += genTextBody(node["p:txBody"], slideLayoutSpNode, slideMasterSpNode, type);
		}
		result += "</div>";
		
	} else {
	
		result += "<div class='block content " + getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) +
				"' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
				"' style='" + 
					getPosition(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) + 
					getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode) + 
					getBorder(node, false) +
					getFill(node, false) +
					" z-index: " + order + ";" +
				"'>";
		
		// TextBody
		if (node["p:txBody"] !== undefined) {
			result += genTextBody(node["p:txBody"], slideLayoutSpNode, slideMasterSpNode, type);
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
		case "emf": // Not native support
			mimeType = "image/emf";
			break;
		case "wmf": // Not native support
			mimeType = "image/wmf";
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
			var tableNode = getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl"]);
			var xfrmNode = getTextByPathList(node, ["p:xfrm"]);
			var tableHtml = "<table style='" + getPosition(xfrmNode, undefined, undefined) + getSize(xfrmNode, undefined, undefined) + "'>";
			
			var trNodes = tableNode["a:tr"];
			if (trNodes.constructor === Array) {
				for (var i=0; i<trNodes.length; i++) {
					tableHtml += "<tr>";
					var tcNodes = trNodes[i]["a:tc"];
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
			result = tableHtml;
			break;
		case "http://schemas.openxmlformats.org/drawingml/2006/chart":
			break;
		case "http://schemas.openxmlformats.org/drawingml/2006/diagram":
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

function genTextBody(textBodyNode, slideLayoutSpNode, slideMasterSpNode, type) {
	
	var text = "";
	
	if (textBodyNode === undefined) {
		return text;
	}

	if (textBodyNode["a:p"].constructor === Array) {
		// multi p
		for (var i=0; i<textBodyNode["a:p"].length; i++) {
			var pNode = textBodyNode["a:p"][i];
			var rNode = pNode["a:r"];
			text += "<div class='" + getHorizontalAlign(pNode, slideLayoutSpNode, slideMasterSpNode, type) + "'>";
			text += genBuChar(pNode);
			if (rNode === undefined) {
				// without r
				text += genSpanElement(pNode, slideLayoutSpNode, slideMasterSpNode, type);
			} else if (rNode.constructor === Array) {
				// with multi r
				for (var j=0; j<rNode.length; j++) {
					text += genSpanElement(rNode[j], slideLayoutSpNode, slideMasterSpNode, type);
				}
			} else {
				// with one r
				text += genSpanElement(rNode, slideLayoutSpNode, slideMasterSpNode, type);
			}
			text += "</div>";
		}
	} else {
		// one p
		var pNode = textBodyNode["a:p"];
		var rNode = pNode["a:r"];
		text += "<div class='" + getHorizontalAlign(pNode, slideLayoutSpNode, slideMasterSpNode, type) + "'>";
		text += genBuChar(pNode);
		if (rNode === undefined) {
			// without r
			text += genSpanElement(pNode, slideLayoutSpNode, slideMasterSpNode, type);
		} else if (rNode.constructor === Array) {
			// with multi r
			for (var j=0; j<rNode.length; j++) {
				text += genSpanElement(rNode[j], slideLayoutSpNode, slideMasterSpNode, type);
			}
		} else {
			// with one r
			text += genSpanElement(rNode, slideLayoutSpNode, slideMasterSpNode, type);
		}
		text += "</div>";
	}
	
	return text;
}

function genBuChar(node) {
	var pPrNode = node["a:pPr"];
	var buChar = getTextByPathList(pPrNode, ["a:buChar", "attrs", "char"]);
	if (buChar !== undefined) {
		var buFontAttrs = getTextByPathList(pPrNode, ["a:buFont", "attrs"]);
		if (buFontAttrs !== undefined) {
			var marginLeft = parseInt( getTextByPathList(pPrNode, ["attrs", "marL"]) ) * 96 / 914400;
			var marginRight = parseInt(buFontAttrs["pitchFamily"]);
			if (isNaN(marginLeft)) {
				marginLeft = 0;
			}
			if (isNaN(marginRight)) {
				marginRight = 0;
			}
			var typeface = buFontAttrs["typeface"];
			
			return "<span style='font-family: " + typeface + 
					"; margin-left: " + marginLeft + "px" +
					"; margin-right: " + marginRight + "pt;'>" + buChar + "</span>";
		}
	}
	
	return "";
}

function genSpanElement(node, slideLayoutSpNode, slideMasterSpNode, type) {
	
	var text = node["a:t"];
	if (typeof text !== 'string') {
		text = getTextByPathList(node, ["a:fld", "a:t"]);
		if (typeof text !== 'string') {
			text = "&nbsp;";
			//debug("XXX: " + JSON.stringify(node));
		}
	}
	
	return "<span class='text-block' style='color: " + getFontColor(node) + 
				"; font-size: " + getFontSize(node, slideLayoutSpNode, slideMasterSpNode, type) + 
				"; font-family: " + getFontType(node) + 
				"; font-weight: " + getFontBold(node) + 
				"; font-style: " + getFontItalic(node) + 
				"; text-decoration: " + getFontDecoration(node) +
				"; vertical-align: " + getTextVerticalAlign(node) + 
				";'>" + text.replace(/\s/i, "&nbsp;") + "</span>";
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

function getHorizontalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) {
	//debug(node);
	var algn = getTextByPathList(node, ["a:pPr", "attrs", "algn"]);
	if (algn === undefined) {
		algn = getTextByPathList(slideLayoutSpNode, ["p:txBody", "a:p", "a:pPr", "attrs", "algn"]);
		if (algn === undefined) {
			algn = getTextByPathList(slideMasterSpNode, ["p:txBody", "a:p", "a:pPr", "attrs", "algn"]);
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

function getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type) {
	
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

function getFontType(node) {
	var typeface = getTextByPathList(node, ["a:rPr", "a:latin", "attrs", "typeface"]);
	return (typeface === undefined) ? "inherit" : typeface;
}

function getFontColor(node) {
	var color = getTextByPathStr(node, "a:rPr a:solidFill a:srgbClr attrs val");
	return (color === undefined) ? "#000" : "#" + color;
}

function getFontSize(node, slideLayoutSpNode, slideMasterSpNode, type) {
	var fontSize = undefined;
	if (node["a:rPr"] !== undefined) {
		fontSize = parseInt(node["a:rPr"]["attrs"]["sz"]) / 100;
	}
	
	if ((isNaN(fontSize) || fontSize === undefined)) {
		var sz = getTextByPathList(slideLayoutSpNode, ["p:txBody", "a:lstStyle", "a:lvl1pPr", "a:defRPr", "attrs", "sz"]);
		fontSize = parseInt(sz) / 100;
	}	
	/*
	if ((isNaN(fontSize) || fontSize === undefined) && slideLayoutSpNode !== undefined && slideLayoutSpNode["a:defRPr"] !== undefined) {
		fontSize = parseInt(slideLayoutSpNode["a:defRPr"]["attrs"]["sz"]) / 100;
	}
	*/
	
	if (isNaN(fontSize)) {
		if (type == "title" || type == "subTitle" || type == "ctrTitle") {
			fontSize = titleFontSize;
		} else if (type === undefined) {
			fontSize = otherFontSize;
		}
	}
	
	var baseline = getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
	if (baseline !== undefined && !isNaN(fontSize)) {
		fontSize -= 10;
	}
	
	return isNaN(fontSize) ? "inherit" : (fontSize + "pt");
}

function getFontBold(node) {
	return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["b"] === "1") ? "bold" : "initial";
}

function getFontItalic(node) {
	return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["i"] === "1") ? "italic" : "normal";
}

function getFontDecoration(node) {
	return (node["a:rPr"] !== undefined && node["a:rPr"]["attrs"]["u"] === "sng") ? "underline" : "initial";
}

function getTextVerticalAlign(node) {
	var baseline = getTextByPathList(node, ["a:rPr", "attrs", "baseline"]);
	if (baseline === undefined) {
		return "";
	} else {
		baseline = parseInt(baseline) / 1000;
		return baseline + "%";
	}
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
	
	// 2. drawingML namespace
	if (borderColor === undefined) {
		var schemeClrNode = getTextByPathList(node, ["p:style", "a:lnRef", "a:schemeClr"]);
		var schemeClr = "a:" + getTextByPathList(schemeClrNode, ["attrs", "val"]);
		var borderColor = getTextByPathList(themeContent, ["a:theme", "a:themeElements", "a:clrScheme", schemeClr, "a:srgbClr", "attrs", "val"]);
		if (borderColor === undefined) {
			borderColor = "FFF";
		}
		
		var shade = getTextByPathList(schemeClrNode, ["a:shade", "attrs", "val"]);
		if (shade !== undefined) {
			shade = parseInt(shade) / 100000;
			var color = new colz.Color("#" + borderColor);
			color.setLum(color.hsl.l * shade);
			borderColor = color.hex.replace("#", "");
		}
		
	}
	borderColor = "#" + borderColor;
	
	// Border type
	var borderType = getTextByPathList(lineNode, ["a:prstDash", "attrs", "val"]);
	switch (borderType) {
		case "solid":
			cssText += "solid";
			break;
		case "dash":
			cssText += "dashed";
			break;
		case "dashDot":
			cssText += "dashed";
			break;
		case "dot":
			cssText += "dotted";
			break;
		case "lgDash":
			cssText += "dashed";
			break;
		case "lgDashDotDot":
			cssText += "dashed";
			break;
		case "sysDash":
			cssText += "dashed";
			break;
		case "sysDashDot":
			cssText += "dashed";
			break;
		case "sysDashDotDot":
			cssText += "dashed";
			break;
		case "sysDot":
			cssText += "dotted";
			break;
		case undefined:
			//console.log(borderType);
		default:
			//console.warn(borderType);
			//cssText += "#000 solid";
	}
	
	if (isSvgMode) {
		return {"color": borderColor, "width": borderWidth, "type": borderType};
	} else {
		return cssText + ";";
	}
}

function getFill(node, isSvgMode) {
	
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
		fillColor = getTextByPathList(themeContent, ["a:theme", "a:themeElements", "a:clrScheme", schemeClr, "a:srgbClr", "attrs", "val"]);
	}
	
	// 2. drawingML namespace
	if (fillColor === undefined) {
		var schemeClr = "a:" + getTextByPathList(node, ["p:style", "a:fillRef", "a:schemeClr", "attrs", "val"]);
		fillColor = getTextByPathList(themeContent, ["a:theme", "a:themeElements", "a:clrScheme", schemeClr, "a:srgbClr", "attrs", "val"]);
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
			return "#FFF";
		} else {
			return "background-color: " + fillColor + ";";
		}
		
	}
	
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
 * eachChild
 * @param {Object} node
 * @param {function} doFunction
 */
function eachChild(node, doFunction) {
	if (node.constructor === Array) {
		var l = node.length;
		for (var i=0; i<l; i++) {
			doFunction(node[i]);
		}
	} else {
		doFunction(node);
	}
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
