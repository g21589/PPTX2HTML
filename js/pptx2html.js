var titleFontSize = 42;
var bodyFontSize = 20;
var otherFontSize = 18;

function openXMLFromZip(zipObj, fiilename) {
	return $($.parseXML(zipObj.file(fiilename).asText()));
}

function escapeHtml(text) {
	var map = {
		'&': '&amp;',
		'<': '&lt;',
		'>': '&gt;',
		'"': '&quot;',
		"'": '&#039;'
	};

	return text.replace(/[&<>"']/g, function(m) { return map[m]; });
}

function getContentTypes(zip) {
	var $contentTypesXML = $($.parseXML(zip.file("[Content_Types].xml").asText()));
	
	var slidesLocArray = [];
	var slides = $contentTypesXML.find(
		"Override[ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"]");
	slides.each(function(index) {
		slidesLocArray.push($(this).attr("PartName").substr(1));
	});
	
	var slideLayoutsLocArray = [];
	var slideLayouts = $contentTypesXML.find(
		"Override[ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"]");
	slideLayouts.each(function(index) {
		slideLayoutsLocArray.push($(this).attr("PartName").substr(1));
	});
	
	return {
		"slides": slidesLocArray,
		"slideLayouts": slideLayoutsLocArray
	};
}

function getSpNodeByID($xml, id) {
	return $xml.find("cNvPr[id=\"" + id + "\"]").parent().parent();
}

function getFontType($slideSpNode) {
	var type = $slideSpNode.find("pPr").attr("typeface");
	if (typeof type == 'undefined') {
		type = $slideSpNode.find("latin").attr("typeface");
	}
	return typeof type != 'undefined' ? type : "inherit";
}

function getFontColor($slideSpNode) {
	var color = $slideSpNode.find("srgbClr").attr("val");
	if (typeof color != 'undefined') {
		color = "#" + color;
	} else {
		color = "#000";
	}
	return color;
}

function getFontSize($slideSpNode, $slideLayoutSpNode, type) {
	var fontSize = (parseInt($slideSpNode.find("rPr").attr("sz")) / 100);
	if (isNaN(fontSize)) {
		fontSize = (parseInt($slideLayoutSpNode.find("defRPr").attr("sz")) / 100);
	}
	if (isNaN(fontSize)) {
		if (type == "title" || type == "subTitle" || type == "ctrTitle") {
			fontSize = titleFontSize;
		} else {
			fontSize = otherFontSize;
		}
	}
	return isNaN(fontSize) ? "inherit" : (fontSize + "pt");
}

function getFontBold($slideSpNode) {
	return $slideSpNode.find("rPr").attr("b") === "1" ? "bold" : "initial";
}

function getFontItalic($slideSpNode) {
	return $slideSpNode.find("rPr").attr("i") === "1" ? "italic" : "normal";
}

function getPosition($slideSpNode, $slideLayoutSpNode, $slideMasterSpNode) {
	var off = $slideSpNode.find("off");	
	var x = parseInt(off.attr("x")) * 96 / 914400;
	var y = parseInt(off.attr("y")) * 96 / 914400;
	if (isNaN(x) || isNaN(y)) {
		// Get info from layoutXML
		off = $slideLayoutSpNode.find("off");
		x = parseInt(off.attr("x")) * 96 / 914400;
		y = parseInt(off.attr("y")) * 96 / 914400;
	}
	if (isNaN(x) || isNaN(y)) {
		// Get info from masterXML
		off = $slideMasterSpNode.find("off");
		x = parseInt(off.attr("x")) * 96 / 914400;
		y = parseInt(off.attr("y")) * 96 / 914400;
	}
	//console.log([x, y]);
	return (isNaN(x) || isNaN(y)) ? "" : "top:" + y + "px; left:" + x + "px;";
}

function getSize($slideSpNode, $slideLayoutSpNode, $slideMasterSpNode) {
	var ext = $slideSpNode.find("ext");
	var w = parseInt(ext.attr("cx")) * 96 / 914400;
	var h = parseInt(ext.attr("cy")) * 96 / 914400;
	if (isNaN(w) || isNaN(h)) {
		// Get info from layoutXML
		ext = $slideLayoutSpNode.find("ext");
		w = parseInt(ext.attr("cx")) * 96 / 914400;
		h = parseInt(ext.attr("cy")) * 96 / 914400;
	}
	if (isNaN(w) || isNaN(h)) {
		// Get info from masterXML
		ext = $slideMasterSpNode.find("ext");
		w = parseInt(ext.attr("cx")) * 96 / 914400;
		h = parseInt(ext.attr("cy")) * 96 / 914400;
		
	}
	//console.log([w, h]);
	return (isNaN(w) || isNaN(h)) ? "" : "width:" + w + "px; height:" + h + "px;";
}

function getSlideSize(zip) {
	// Pixel = EMUs * Resolution / 914400;  (Resolution = 96)
	var $presentationXML = $($.parseXML(zip.file("ppt/presentation.xml").asText()));
	var sizeNode = $presentationXML.find("sldSz");
	return {
		"width": (parseInt(sizeNode.attr("cx")) * 96 / 914400),
		"height": (parseInt(sizeNode.attr("cy")) * 96 / 914400)
	};
}

function base64ArrayBuffer(arrayBuffer) {
	var base64    = '';
	var encodings = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
	var bytes         = new Uint8Array(arrayBuffer);
	var byteLength    = bytes.byteLength;
	var byteRemainder = byteLength % 3;
	var mainLength    = byteLength - byteRemainder;

	var a, b, c, d;
	var chunk;

	for (var i = 0; i < mainLength; i = i + 3) {
		chunk = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
		a = (chunk & 16515072) >> 18;
		b = (chunk & 258048)   >> 12;
		c = (chunk & 4032)     >>  6;
		d = chunk & 63;
		base64 += encodings[a] + encodings[b] + encodings[c] + encodings[d];
	}

	if (byteRemainder == 1) {
		chunk = bytes[mainLength];
		a = (chunk & 252) >> 2;
		b = (chunk & 3)   << 4;
		base64 += encodings[a] + encodings[b] + '==';
	} else if (byteRemainder == 2) {
		chunk = (bytes[mainLength] << 8) | bytes[mainLength + 1];
		a = (chunk & 64512) >> 10;
		b = (chunk & 1008)  >>  4;
		c = (chunk & 15)    <<  2;
		base64 += encodings[a] + encodings[b] + encodings[c] + '=';
	}

	return base64;
}

(function () {

	if (!window.FileReader || !window.ArrayBuffer) {
		$("#error_block").removeClass("hidden").addClass("show");
		return;
	}

	var $result = $("#result");
	$("#file").on("change", function(evt) {

		$result.html("");
		$("#result_block").removeClass("hidden").addClass("show");

		var files = evt.target.files;
		for (var i = 0, f; f = files[i]; i++) {

			var reader = new FileReader();

			// Closure to capture the file information.
			reader.onload = (function(theFile) {
				return function(e) {
					var $title = $("<h4>", {
						text : theFile.name
					});
					$result.append($title);
					var $fileContent = $("<ul>");
					try {

						var dateBefore = new Date();
						// read the content of the file with JSZip
						var zip = new JSZip(e.target.result);
						var dateAfter = new Date();

						$title.append($("<span>", {
							text:" (parsed in " + (dateAfter - dateBefore) + "ms)"
						}));
						
						// Get files information in the pptx
						var filesInfo = getContentTypes(zip);
						
						// Size information
						var slideSize = getSlideSize(zip);
						
						// Open each slide XML
						$.each(filesInfo["slides"], function (index, name) {
							
							console.log(name);
							var context = "";
							
							var slideXMLText = zip.file(name).asText();
							var $slideXML = $($.parseXML(slideXMLText));
							
							// Read relationship filename of the slide (Get slideLayoutXX.xml)
							// @name: ppt/slides/slide1.xml
							// @resName: ppt/slides/_rels/slide1.xml.rels
							var resName = name.replace("slides/slide", "slides/_rels/slide") + ".rels";
							var $resTarget = openXMLFromZip(zip, resName)
								.find("Relationship[Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\"]")
								.attr("Target")
								.replace("../", "ppt/");
							console.log($resTarget);
							
							// Open slideLayoutXX.xml
							var $slideLayoutXML = openXMLFromZip(zip, $resTarget);
							
							// Read slide master filename of the slidelayout (Get slideMasterXX.xml)
							// @resName: ppt/slideLayouts/slideLayout1.xml
							// @masterName: ppt/slideLayouts/_rels/slideLayout1.xml.rels
							var masterName = $resTarget.replace("slideLayouts/slideLayout", "slideLayouts/_rels/slideLayout") + ".rels";
							var $masterTarget = openXMLFromZip(zip, masterName)
								.find("Relationship[Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\"]")
								.attr("Target")
								.replace("../", "ppt/");
							console.log($masterTarget);
							
							// Open slideMasterXX.xml
							var $slideMasterXML = openXMLFromZip(zip, $masterTarget);
							
							/* 
							 * Process Slide Master
							 *   titleStyle
							 *   bodyStyle
							 *   otherStyle
							 */
							var $titleStyleNode = $slideMasterXML.find("titleStyle");
							var $bodyStyleNode = $slideMasterXML.find("bodyStyle");
							var $otherStyleNode = $slideMasterXML.find("otherStyle");
							
							titleFontSize = parseInt($titleStyleNode.find("defRPr").attr("sz")) / 100;
							bodyFontSize = parseInt($bodyStyleNode.find("defRPr").attr("sz")) / 100;   // TODO: level
							otherFontSize = parseInt($otherStyleNode.find("defRPr").attr("sz")) / 100; // TODO: level
							
							// Parse the slide context and rander into html
							$slideXML.find("sp").each(function(index, slideSpNode) {
								var $slideSpNode = $(slideSpNode);
								var type = $slideSpNode.find("ph").attr("type");
								var text = $slideSpNode.find("t").text();
								var id = $slideSpNode.find("cNvPr").attr("id");
								console.log("  id: " + id);
								
								var $slideLayoutSpNode = getSpNodeByID($slideLayoutXML, id);
								var $slideMasterSpNode = getSpNodeByID($slideMasterXML, id);
								/*
								if (type == "title") {
									text = "<div class='block title'><h2 data-toggle='tooltip' data-placement='top' title='title'>" + text + "</h2></div>";
								} else if (type == "subTitle") {
									text = "<div class='block subTitle'><h2 data-toggle='tooltip' data-placement='top' title='title'>" + text + "</h2></div>";
								}else if (type == "ctrTitle") {
									text = "<div class='block ctrTitle'><h1 data-toggle='tooltip' data-placement='top' title='ctrTitle'>" + text + "</h1></div>";
								} else if (type == "dt") {
									//text = "<div data-toggle='tooltip' data-placement='top' title='dt'>" + text + "</div>";
									text = "";
								} else if (type == "sldNum") {
									//text = "<div data-toggle='tooltip' data-placement='top' title='sldNum'>" + text + "</div>";
									text = "";
								} else {
								*/
									text = "<div class='block content' style='" + getPosition($slideSpNode, $slideLayoutSpNode, $slideMasterSpNode) + getSize($slideSpNode, $slideLayoutSpNode, $slideMasterSpNode) + "'>";
									$slideSpNode.find("p").each(function(index, node) {
										var $node = $(node);
										text += "<div style='color: " + getFontColor($node) + 
												"; font-size: " + getFontSize($node, $slideLayoutSpNode, type) + 
												"; font-weight: " + getFontBold($node) + 
												"; font-style: " + getFontItalic($node) + 
												"; font-family: " + getFontType($node) + 
												";'>" + $node.find("t").text() + "</div>";
									});
									text += "</div>";
								//}
								context += text;
							});
							
							$slideXML.find("pic").each(function(index, node) {
								var $node = $(node);
								var rid = $node.find("blip").attr("r:embed");
								var $xfrmNode = $node.find("spPr").find("xfrm");
								var imgName = openXMLFromZip(zip, resName).find("Relationship[Id=\"" + rid + "\"]").attr("Target").replace("../", "ppt/");
								var imgArrayBuffer = zip.file(imgName).asArrayBuffer();
								context += "<div class='block content' style='" + getPosition($xfrmNode, null, null) + getSize($xfrmNode, null, null) + 
										   "'><img src=\"data:image/jpeg;base64," + base64ArrayBuffer(imgArrayBuffer) + "\" style='width: 100%; height: 100%'/></div>";
							});
							
							$fileContent.append($("<li>", {
								"class" : "slide",
								html : name + 
									   "<section style='width:" + slideSize.width + "px; height:" + slideSize.height + "px;'>" + context + "</section>" +
									   "<details><pre><code class='xml'>" + escapeHtml(vkbeautify.xml(slideXMLText, 4)) + "</code></pre></details>"
							}));
							
						});
						
					} catch(e) {
						$fileContent = $("<div>", {
							"class" : "alert alert-danger",
							text : "Error reading " + theFile.name + " : " + e.message
						});
					}
					
					$result.append($fileContent);
					
					$('pre code').each(function(i, block) {
						hljs.highlightBlock(block);
					});
					
				}
			})(f);

			// Read the file
			reader.readAsArrayBuffer(f);
		}
	});
})();
