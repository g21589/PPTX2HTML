importScripts(
	'./jszip.min.js',
	'./highlight.min.js',
	'./colz.class.min.js',
	'./highlight.min.js',
	'./functions.js',
	'./tXmlUnfolded.js'
);

onmessage = function(e) {
	
	var zip = new JSZip(e.data);
	
	var pptxThumbImg = base64ArrayBuffer(zip.file("docProps/thumbnail.jpeg").asArrayBuffer());
	self.postMessage({
		"type": "pptx-thumb",
		"data": pptxThumbImg
	});
	
	var filesInfo = getContentTypes(zip);
	self.postMessage({
		"type": "INFO",
		"data": JSON.stringify( filesInfo )
	});
	
	var slideSize = getSlideSize(zip);
	self.postMessage({
		"type": "INFO",
		"data": JSON.stringify( slideSize )
	});
	
	for (var i=0; i<filesInfo["slides"].length; i++) {
		var filename = filesInfo["slides"][i];
		processSingleSlide(zip, filename, i);
		//self.postMessage({
		//	"type": "INFO",
		//	"data": JSON.stringify( content )
		//});
	}
}

function readXmlFile(zip, filename) {
	return tXml(zip.file(filename).asText());
}

function getContentTypes(zip) {
	var ContentTypesJson = readXmlFile(zip, "[Content_Types].xml");
	var subObj = ContentTypesJson["?xml"]["Types"]["Override"];
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
	var sldSzAttrs = content["?xml"]["p:presentation"]["p:sldSz"]["attrs"]
	return {
		"width": parseInt(sldSzAttrs["cx"]) * 96 / 914400,
		"height": parseInt(sldSzAttrs["cy"]) * 96 / 914400
	};
}

function processSingleSlide(zip, sldFileName, index) {
	
	// =====< Step 1 >=====
	// Read relationship filename of the slide (Get slideLayoutXX.xml)
	// @sldFileName: ppt/slides/slide1.xml
	// @resName: ppt/slides/_rels/slide1.xml.rels
	resName = sldFileName.replace("slides/slide", "slides/_rels/slide") + ".rels";
	var resContent = readXmlFile(zip, resName);
	
	self.postMessage({
		"type": "INFO",
		"data": JSON.stringify( resContent["?xml"]["Relationships"]["Relationship"] )
	});

		//.find("Relationship[Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\"]")
		//.attr("Target")
		//.replace("../", "ppt/");
	
	// =====< Step 2 >=====
	// Read slide master filename of the slidelayout (Get slideMasterXX.xml)
	// @resName: ppt/slideLayouts/slideLayout1.xml
	// @masterName: ppt/slideLayouts/_rels/slideLayout1.xml.rels
	/*
	var masterName = $resTarget.replace("slideLayouts/slideLayout", "slideLayouts/_rels/slideLayout") + ".rels";
	var $masterTarget = openXMLFromZip(zip, masterName)
		.find("Relationship[Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\"]")
		.attr("Target")
		.replace("../", "ppt/");
	console.log($masterTarget);
	
	// Open slideMasterXX.xml
	$slideMasterXML = openXMLFromZip(zip, $masterTarget);
	*/
	
	// =====< Step 3 >=====
	var content = readXmlFile(zip, sldFileName);
	var nodes = content["?xml"]["p:sld"]["p:cSld"]["p:spTree"];
	for (var nodeKey in nodes) {
		processNodesInSlide(nodeKey, nodes[nodeKey]);
	}
	
}

function processNodesInSlide(nodeKey, nodeValue) {
	
	var result = "";
	
	switch (nodeKey) {
		case "p:sp":	// Shape, Text
			//result += processSpNode($node, $slideLayoutXML, $slideMasterXML);
			break;
		case "p:pic":	// Picture
			result += processPicNode(nodeValue, null);
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
			//result += "<div class='block group'>";
			//$node.children().each(processNodesInSlide);
			//result += "</div>";
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

function processPicNode($node, resName) {
	
	/*
	var rid = $node.find("blip").attr("r:embed");
	var $xfrmNode = $node.find("spPr").find("xfrm");
	var imgName = openXMLFromZip(zip, resName).find("Relationship[Id=\"" + rid + "\"]").attr("Target").replace("../", "ppt/");
	var imgFileExt = extractFileExtension(imgName).toLowerCase();
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
	return "<div class='block content' style='" + getPosition($xfrmNode, null, null) + getSize($xfrmNode, null, null) + 
			   "'><img src=\"data:" + mimeType + ";base64," + base64ArrayBuffer(imgArrayBuffer) + "\" style='width: 100%; height: 100%'/></div>";
	*/
	
	return "";
}
