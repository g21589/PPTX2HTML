importScripts(
	'./jszip.min.js',
	'./highlight.min.js',
	'./colz.class.min.js',
	'./highlight.min.js',
	'./tXml.js',
	'./functions.js'
);

onmessage = function(e) {
	
	var dateBefore = new Date();
	
	var zip = new JSZip(e.data);
	
	var pptxThumbImg = base64ArrayBuffer(zip.file("docProps/thumbnail.jpeg").asArrayBuffer());
	self.postMessage({
		"type": "pptx-thumb",
		"data": pptxThumbImg
	});
	
	var filesInfo = getContentTypes(zip);
	/*
	self.postMessage({
		"type": "INFO",
		"data": JSON.stringify( filesInfo )
	});
	*/
	
	var slideSize = getSlideSize(zip);
	/*
	self.postMessage({
		"type": "INFO",
		"data": JSON.stringify( slideSize )
	});
	*/
	
	for (var i=0; i<filesInfo["slides"].length; i++) {
		var filename = filesInfo["slides"][i];
		var slideHtml = processSingleSlide(zip, filename, i, slideSize);
		self.postMessage({
			"type": "slide",
			"data": slideHtml
		});
		self.postMessage({
			"type": "progress-update",
			"data": (i + 1) * 100 / filesInfo["slides"].length
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

function processSingleSlide(zip, sldFileName, index, slideSize) {
	
	self.postMessage({
		"type": "INFO",
		"data": "Processing slide" + index
	});
	
	// =====< Step 1 >=====
	// Read relationship filename of the slide (Get slideLayoutXX.xml)
	// @sldFileName: ppt/slides/slide1.xml
	// @resName: ppt/slides/_rels/slide1.xml.rels
	resName = sldFileName.replace("slides/slide", "slides/_rels/slide") + ".rels";
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
				result += processNodesInSlide(nodeKey, nodes[nodeKey][i], warpObj);
			}
		} else {
			result += processNodesInSlide(nodeKey, nodes[nodeKey], warpObj);
		}
	}
	
	return result + "</section></li>";
}

function indexNodes(content) {
	
	var keys = Object.keys(content);
	var spTreeNode = content[keys[0]]["p:cSld"]["p:spTree"];
	
	var idTable = {};
	var idxTable = {};
	
	for (var key in spTreeNode) {

		if (key == "p:nvGrpSpPr" || key == "p:grpSpPr") {
			continue;
		}
		
		var targetNode = spTreeNode[key];
		
		if (targetNode.constructor === Array) {
			for (var i=0; i<targetNode.length; i++) {
				var id = (targetNode[i]["p:nvSpPr"]["p:cNvPr"] === undefined) ? undefined : targetNode[i]["p:nvSpPr"]["p:cNvPr"]["attrs"]["id"];
				var idx = (targetNode[i]["p:nvSpPr"]["p:nvPr"] === undefined || targetNode[i]["p:nvSpPr"]["p:nvPr"]["p:ph"] === undefined) ? 
					undefined : targetNode[i]["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["idx"];
				
				if (id !== undefined) {
					idTable[id] = targetNode[i];
				}
				if (idx !== undefined) {
					idxTable[idx] = targetNode[i];
				}
			}
		} else {
			var id = (targetNode["p:nvSpPr"]["p:cNvPr"] === undefined) ? undefined : targetNode["p:nvSpPr"]["p:cNvPr"]["attrs"]["id"];
			var idx = (targetNode["p:nvSpPr"]["p:nvPr"] === undefined || targetNode["p:nvSpPr"]["p:nvPr"]["p:ph"] === undefined) ? 
				undefined : targetNode["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["idx"];
			
			if (id !== undefined) {
				idTable[id] = targetNode;
			}
			if (idx !== undefined) {
				idxTable[idx] = targetNode;
			}
		}
		
	}
	
	return {idTable, idxTable};
}

function processNodesInSlide(nodeKey, nodeValue, warpObj) {
	
	var result = "";
	
	switch (nodeKey) {
		case "p:sp":	// Shape, Text
			result += processSpNode(nodeValue, warpObj);
			break;
		case "p:pic":	// Picture
			result += processPicNode(nodeValue, warpObj);
			break;
		case "p:graphicFrame":	// Chart, Diagram, Table
			/*
			if ($node.find("graphicData").attr("uri") === 
					"http://schemas.openxmlformats.org/drawingml/2006/table") {
				// Table
				$tableNode = $node.find("graphic").find("tbl");
				$xfrmNode = $node.find("xfrm");
				var tableHtml = "<table style='" + getPosition($xfrmNode, null, null) + getSize($xfrmNode, null, null) + "'>";
				$tableNode.find("tr").each(function(index, node) {
					var $node = $(node);
					tableHtml += "<tr>";
					$node.find("tc").each(function(index, node) {
						var $node = $(node);
						tableHtml += "<td>" + $node.find("t").text() + "</td>";
					});
					tableHtml += "</tr>";
				});
				tableHtml += "</table>";
				result += tableHtml;
			} else if ($node.find("graphicData").attr("uri") === 
					"http://schemas.openxmlformats.org/drawingml/2006/chart") {
				// TODO: Chart
				
			} else {
				// TODO: Diagram
				
			}
			*/
			break;
		case "p:grpSp":	// 群組
			result += "<div class='block group'>";			
			for (var nodeKey in nodeValue) {
				if (nodeValue[nodeKey].constructor === Array) {
					for (var i=0; i<nodeValue[nodeKey].length; i++) {
						result += processNodesInSlide(nodeKey, nodeValue[nodeKey][i], warpObj);
					}
				} else {
					result += processNodesInSlide(nodeKey, nodeValue[nodeKey], warpObj);
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
			/*
			var $xfrmNode = $node.find("xfrm");
			var x = parseInt($xfrmNode.find("off").attr("x")) * 96 / 914400;
			var y = parseInt($xfrmNode.find("off").attr("y")) * 96 / 914400;
			var chx = parseInt($xfrmNode.find("chOff").attr("x")) * 96 / 914400;
			var chy = parseInt($xfrmNode.find("chOff").attr("y")) * 96 / 914400;
			var cx = parseInt($xfrmNode.find("ext").attr("cx")) * 96 / 914400;
			var cy = parseInt($xfrmNode.find("ext").attr("cy")) * 96 / 914400;
			var chcx = parseInt($xfrmNode.find("chExt").attr("cx")) * 96 / 914400;
			var chcy = parseInt($xfrmNode.find("chExt").attr("cy")) * 96 / 914400;
			result = result.replace(new RegExp('>$'), " style='top: " + (y - chy) + "px; left: " + (x - chx) + 
						"px; width: " + cx + "px; height: " + cy + "px;'>");
			*/
			break;
		default:
	}
	
	return result;
	
}

function processSpNode(node, warpObj) {
	
	var id = node["p:nvSpPr"]["p:cNvSpPr"]["attrs"]["id"];
	var name = node["p:nvSpPr"]["p:cNvSpPr"]["attrs"]["name"];
	var idx = (node["p:nvSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["idx"];
	var type = (node["p:nvSpPr"]["p:nvPr"]["p:ph"] === undefined) ? undefined : node["p:nvSpPr"]["p:nvPr"]["p:ph"]["attrs"]["type"];
	
	var slideLayoutSpNode = undefined;
	var slideMasterSpNode = undefined;
	
	if (idx === undefined) {
		slideLayoutSpNode = warpObj["slideLayoutTables"]["idTable"][id];
		slideMasterSpNode = warpObj["slideMasterTables"]["idTable"][id];
	} else {
		var _id = warpObj["slideLayoutTables"]["idxTable"][idx]["p:nvSpPr"]["p:cNvPr"]["attrs"]["id"];
		slideMasterSpNode = warpObj["slideMasterTables"]["idTable"][_id];
	}
	
	debug( {id, name, idx, type} );
	//debug( JSON.stringify( node ) );
	
	var text = "";
	
	text += "<div class='block content " + getAlign(node, slideLayoutSpNode, slideMasterSpNode, null) +
			"' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
			"' style='" + 
				getPosition(node, slideLayoutSpNode, slideMasterSpNode) + 
				getSize(node, slideLayoutSpNode, slideMasterSpNode) + 
				//getBorder($node) +
				//getFill($node) +
			"'>";
	
	// Text
	if (node["p:txBody"] === undefined) {
		return text;
	}
	
	var textBodyNode = node["p:txBody"];
	if (textBodyNode["a:p"].constructor === Array) {
		// multi p
		for (var i=0; i<textBodyNode["a:p"].length; i++) {
			var pNode = textBodyNode["a:p"][i];
			var rNode = pNode["a:r"];
			text += "<div>";
			if (rNode === undefined) {
				// without r
				text += genSpanElement(pNode, pNode["a:t"], slideLayoutSpNode, slideMasterSpNode);
			} else if (rNode.constructor === Array) {
				// with multi r
				for (var j=0; j<rNode.length; j++) {
					text += genSpanElement(rNode[j], rNode[j]["a:t"], slideLayoutSpNode, slideMasterSpNode);
				}
			} else {
				// with one r
				text += genSpanElement(rNode, rNode["a:t"], slideLayoutSpNode, slideMasterSpNode);
			}
			text += "</div>";
		}
	} else {
		// one p
		var pNode = textBodyNode["a:p"];
		var rNode = pNode["a:r"];
		text += "<div>";
		if (rNode === undefined) {
			// without r
			text += genSpanElement(pNode, pNode["a:t"], slideLayoutSpNode, slideMasterSpNode);
		} else if (rNode.constructor === Array) {
			// with multi r
			for (var j=0; j<rNode.length; j++) {
				text += genSpanElement(rNode[j], rNode[j]["a:t"], slideLayoutSpNode, slideMasterSpNode);
			}
		} else {
			// with one r
			text += genSpanElement(rNode, rNode["a:t"], slideLayoutSpNode, slideMasterSpNode);
		}
		text += "</div>";
	}
	
/*
	var svgMode = false;
	if ($node.find("style").length > 0) {
		svgMode = true;
		//text = "<svg _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
		//		"' style='" + 
		//			getPosition($node, $slideLayoutSpNode, $slideMasterSpNode) + 
		//			getSize($node, $slideLayoutSpNode, $slideMasterSpNode) + 
					//getBorder($node) +
					//getFill($node) +
		//		"'><ellipse cx='200' cy='80' rx='100' ry='50' style='fill:yellow; stroke:purple; stroke-width:2'/></svg>";
		
		var off = $node.find("off");	
		var x = parseInt(off.attr("x")) * 96 / 914400;
		var y = parseInt(off.attr("y")) * 96 / 914400;
		
		var ext = $node.find("ext");
		var w = parseInt(ext.attr("cx")) * 96 / 914400;
		var h = parseInt(ext.attr("cy")) * 96 / 914400;
		
		var svgDom = $(document.createElement('svg'));
		svgDom.addClass("drawing");
		svgDom.attr({
			"_id": id,
			"_idx": idx,
			"_type": type,
			"_name": name
		});
		svgDom.css({
			"top": y,
			"left": x,
			"width": w,
			"height": h
		});
		
		var fillColor = "#" + $themeXML.find($node.find("style").find("fillRef").find("schemeClr").attr("val")).find("srgbClr").attr("val");
		
		var borderColorStr = $themeXML.find($node.find("style").find("lnRef").find("schemeClr").attr("val")).find("srgbClr").attr("val");
		var borderColor = undefined;
		if (borderColorStr !== undefined) {
			borderColor = new colz.Color("#" + borderColorStr);
			borderColor.setLum(borderColor.hsl.l / 1.5);
		}
		
		switch ($node.find("prstGeom").attr("prst")) {
			case "rect":
				var gDom = $(document.createElement('rect'));
				gDom.attr({
					"x": 0,
					"y": 0,
					"width": w,
					"height": h,
					"fill": fillColor,
					"stroke": (borderColor === undefined || borderColor.rgb === undefined) ? "" : borderColor.rgb.toString(),
					"stroke-width": "1pt"
				});
				svgDom.append(gDom);
				break;
			case "ellipse":
				var gDom = $(document.createElement('ellipse'));
				gDom.attr({
					"cx": w / 2,
					"cy": h / 2,
					"rx": w / 2,
					"ry": h / 2,
					"fill": fillColor,
					"stroke": (borderColor === undefined || borderColor.rgb === undefined) ? "" : borderColor.rgb.toString(),
					"stroke-width": "1pt"
				});
				svgDom.append(gDom);
				break;
			case "roundRect":
				var gDom = $(document.createElement('rect'));
				gDom.attr({
					"x": 0,
					"y": 0,
					"width": w,
					"height": h,
					"rx": 15,
					"ry": 15,
					"fill": fillColor,
					"stroke": (borderColor === undefined || borderColor.rgb === undefined) ? "" : borderColor.rgb.toString(),
					"stroke-width": "1pt"
				});
				svgDom.append(gDom);
				break;
			default:
		}
		text = svgDom[0].outerHTML;
	} else {
		text = "<div class='block content " + getAlign($node, $slideLayoutSpNode, $slideMasterSpNode, type) +
			"' _id='" + id + "' _idx='" + idx + "' _type='" + type + "' _name='" + name +
			"' style='" + 
				getPosition($node, $slideLayoutSpNode, $slideMasterSpNode) + 
				getSize($node, $slideLayoutSpNode, $slideMasterSpNode) + 
				getBorder($node) +
				getFill($node) +
			"'>";
	}
	
	var nodeArr = $node.find("txBody").find("p").each(function(index, node) {
		var $node = $(node);
		text += "<div>";
		var buChar = $node.find("pPr").find("buChar").attr("char");
		if (buChar !== undefined) {
			var $buFontNode = $node.find("pPr").find("buFont");
			var marginLeft = parseInt($node.find("pPr").attr("marL")) * 96 / 914400;
			if (isNaN(marginLeft)) {
				marginLeft = 0;
			}
			var marginRight = parseInt($buFontNode.attr("pitchFamily"));
			if (isNaN(marginRight)) {
				marginRight = 0;
			}
			text += "<span style='font-family: " + $buFontNode.attr("typeface") + 
					"; margin-left: " + marginLeft + "px" +
					"; margin-right: " + marginRight + "pt;'>" + buChar + "</span>";
		}
		
		// With "r"
		var nodeArr = $node.find("r");
		for (var i=0; i<nodeArr.length; i++) {
			var $node = $(nodeArr[i]);
			text += "<span style='color: " + getFontColor($node) + 
					"; font-size: " + getFontSize($node, $slideLayoutSpNode, type) + 
					"; font-weight: " + getFontBold($node) + 
					"; font-style: " + getFontItalic($node) + 
					"; font-family: " + getFontType($node) + 
					"; text-decoration: " + getFontDecoration($node) + 
					";'>" + $node.find("t").text() + "</span>";
		}
		
		// Without "r"
		if (nodeArr.length <= 0) {
			text += "<span style='color: " + getFontColor($node) + 
					"; font-size: " + getFontSize($node, $slideLayoutSpNode, type) + 
					"; font-weight: " + getFontBold($node) + 
					"; font-style: " + getFontItalic($node) + 
					"; font-family: " + getFontType($node) + 
					"; text-decoration: " + getFontDecoration($node) + 
					";'>" + $node.find("t").text() + "</span>";
		}
		
		text += "</div>"
	});
	
	if (!svgMode) {
		text += "</div>";
	}
*/
	text += "</div>";
	return text;
}

function processPicNode(node, warpObj) {
	
	//debug( JSON.stringify( node ) );
	
	var rid = node["p:blipFill"]["a:blip"]["attrs"]["r:embed"];
	var imgName = warpObj["slideResObj"][rid]["target"];
	var imgFileExt = extractFileExtension(imgName).toLowerCase();
	var zip = warpObj["zip"];
	var imgArrayBuffer = zip.file(imgName).asArrayBuffer();
	var mimeType = "";
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
	return "<div class='block content' style='" + getPosition(node, null, null) + getSize(node, null, null) +
			   "'><img src=\"data:" + mimeType + ";base64," + base64ArrayBuffer(imgArrayBuffer) + "\" style='width: 100%; height: 100%'/></div>";
}

function genSpanElement(node, text, slideLayoutSpNode, slideMasterSpNode) {
	return "<span style='color: " + getFontColor(node) + 
				"; font-size: " + getFontSize(node, slideLayoutSpNode, null) + 
				"; font-family: " + getFontType(node) + 
				"; font-weight: " + getFontBold(node) + 
				"; font-style: " + getFontItalic(node) + 
				"; text-decoration: " + getFontDecoration(node) + 
				";'>" + text + "</span>";
}

function getPosition(slideSpNode, slideLayoutSpNode, slideMasterSpNode) {
	
	//debug(JSON.stringify(slideLayoutSpNode));
	//debug(JSON.stringify(slideMasterSpNode));
	
	var off = undefined;
	var x = -1, y = -1;
	
	if (slideSpNode["p:spPr"]["a:xfrm"] !== undefined) {
		off = slideSpNode["p:spPr"]["a:xfrm"]["a:off"]["attrs"];
	} else if (slideLayoutSpNode !== undefined && slideLayoutSpNode["p:spPr"]["a:xfrm"] !== undefined) {
		off = slideLayoutSpNode["p:spPr"]["a:xfrm"]["a:off"]["attrs"];
	} else if (slideMasterSpNode !== undefined && slideMasterSpNode["p:spPr"]["a:xfrm"] !== undefined) {
		off = slideMasterSpNode["p:spPr"]["a:xfrm"]["a:off"]["attrs"];
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
	
	if (slideSpNode["p:spPr"]["a:xfrm"] !== undefined) {
		ext = slideSpNode["p:spPr"]["a:xfrm"]["a:ext"]["attrs"];
	} else if (slideLayoutSpNode !== undefined && slideLayoutSpNode["p:spPr"]["a:xfrm"] !== undefined) {
		ext = slideLayoutSpNode["p:spPr"]["a:xfrm"]["a:ext"]["attrs"];
	} else if (slideMasterSpNode !== undefined && slideMasterSpNode["p:spPr"]["a:xfrm"] !== undefined) {
		ext = slideMasterSpNode["p:spPr"]["a:xfrm"]["a:ext"]["attrs"];
	}
	
	if (ext === undefined) {
		return "";
	} else {
		w = parseInt(ext["cx"]) * 96 / 914400;
		h = parseInt(ext["cy"]) * 96 / 914400;
		return (isNaN(w) || isNaN(h)) ? "" : "width:" + w + "px; height:" + h + "px;";
	}	
	
}

function getAlign(node, slideLayoutSpNode, slideMasterSpNode, type) {
	
	// 上中下對齊: X, <a:bodyPr anchor="ctr">, <a:bodyPr anchor="b">
	var anchor = (node["p:txBody"] === undefined || node["p:txBody"]["a:bodyPr"]["attrs"] === undefined) ? "" : node["p:txBody"]["a:bodyPr"]["attrs"]["anchor"];
	if (anchor === undefined && 
		slideLayoutSpNode !== undefined &&
		slideLayoutSpNode["p:txBody"] !== undefined &&
		slideLayoutSpNode["p:txBody"]["a:bodyPr"]["attrs"] !== undefined) {
		anchor = slideLayoutSpNode["p:txBody"]["a:bodyPr"]["attrs"]["anchor"];
	}
	if (anchor === undefined &&
		slideMasterSpNode !== undefined && 
		slideMasterSpNode["p:txBody"] !== undefined &&
		slideMasterSpNode["p:txBody"]["a:bodyPr"]["attrs"] !== undefined) {
		anchor = slideMasterSpNode["p:txBody"]["a:bodyPr"]["attrs"]["anchor"];
	}
	
	// 左中右對齊: X, <a:pPr algn="ctr"/>, <a:pPr algn="r"/>
	// TODO: in p r
	var algn = "ctr";
	/*
	var algn = node.find("pPr").attr("algn");
	if (algn === undefined) {
		algn = slideLayoutSpNode.find("pPr").attr("algn");
	}
	if (algn === undefined) {
		algn = slideMasterSpNode.find("pPr").attr("algn");
		if (algn === undefined) {
			algn = slideMasterSpNode.find("lvl1pPr").attr("algn");
		}
	}
	*/
	
	/*
	if (type == "title" || type == "subTitle" || type == "ctrTitle") {
		return "center-center";
	}
	*/
	
	if (anchor === "ctr") {
		if (algn === "ctr") {
			return "center-center";
		} else if (algn === "r") {
			return "center-right";
		} else {
			return "center-left";
		}
	} else if (anchor === "b") {
		if (algn === "ctr") {
			return "down-center";
		} else if (algn === "r") {
			return "down-right";
		} else {
			return "down-left";
		}
	} else {
		if (algn === "ctr") {
			return "up-center";
		} else if (algn === "r") {
			return "up-right";
		} else {
			return "up-left";
		}
	}

}

function getFontType(node) {
	return (node["a:rPr"] !== undefined && node["a:rPr"]["a:latin"] !== undefined) ? 
				node["a:rPr"]["a:latin"]["attrs"]["typeface"] : "inherit";
}

function getFontColor(node) {
	var color = undefined;
	if (node["a:rPr"] !== undefined && node["a:rPr"]["a:solidFill"] !== undefined && node["a:rPr"]["a:solidFill"]["a:srgbClr"] !== undefined) {
		color = node["a:rPr"]["a:solidFill"]["a:srgbClr"]["attrs"]["val"];
	}
	if (color === undefined) {
		color = "#" + color;
	} else {
		color = "#000";
	}
	return color;
}


function getFontSize(node, slideLayoutSpNode, type) {
	var fontSize = undefined;
	if (node["a:rPr"] !== undefined) {
		fontSize = parseInt(node["a:rPr"]["attrs"]["sz"]) / 100;
	}
	if (isNaN(fontSize) && slideLayoutSpNode !== undefined && slideLayoutSpNode["a:defRPr"] !== undefined) {
		fontSize = parseInt(slideLayoutSpNode["a:defRPr"]["attrs"]["sz"]) / 100;
	}
	/*
	if (isNaN(fontSize)) {
		if (type == "title" || type == "subTitle" || type == "ctrTitle") {
			fontSize = titleFontSize;
		} else {
			fontSize = otherFontSize;
		}
	}
	*/
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

function debug(data) {
	self.postMessage({"type": "DEBUG", "data": data});
}
