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

function getFontType(element) {
	var type = $(element).find("pPr").attr("typeface");
	if (typeof type == 'undefined') {
		type = $(element).find("latin").attr("typeface");
	}
	return typeof type != 'undefined' ? type : "inherit";
}

function getFontColor(element) {
	var color = $(element).find("srgbClr").attr("val");
	if (typeof color != 'undefined') {
		color = "#" + color;
	} else {
		color = "#000";
	}
	return color;
}

function getFontSize(element, id, $layoutXML) {
	var fontSize = (parseInt($(element).find("rPr").attr("sz")) / 100);
	if (isNaN(fontSize)) {
		fontSize = (parseInt($layoutXML.find("cNvPr[id=\"" + id + "\"]").parent().parent().find("defRPr").attr("sz")) / 100);
	}
	return isNaN(fontSize) ? "inherit" : (fontSize + "pt");
}

function getFontBold(element) {
	return $(element).find("rPr").attr("b") === "1" ? "bold" : "initial";
}

function getFontItalic(element) {
	return $(element).find("rPr").attr("i") === "1" ? "italic" : "normal";
}

function getPosition(element, id, $layoutXML) {
	var off = $(element).find("off");	
	var x = parseInt(off.attr("x")) * 96 / 914400;
	var y = parseInt(off.attr("y")) * 96 / 914400;
	if (isNaN(x) || isNaN(y)) {
		// Get info from layoutXML
		off = $layoutXML.find("cNvPr[id=\"" + id + "\"]").parent().parent().find("off");
		x = parseInt(off.attr("x")) * 96 / 914400;
		y = parseInt(off.attr("y")) * 96 / 914400;
		console.log([x, y]);
	}
	return (isNaN(x) || isNaN(y)) ? "" : "top:" + y + "px; left:" + x + "px;";
}

function getSize(element, id, $layoutXML) {
	var ext = $(element).find("ext");
	var w = parseInt(ext.attr("cx")) * 96 / 914400;
	var h = parseInt(ext.attr("cy")) * 96 / 914400;
	if (isNaN(w) || isNaN(h)) {
		// Get info from layoutXML
		ext = $layoutXML.find("cNvPr[id=\"" + id + "\"]").parent().parent().find("ext");
		w = parseInt(ext.attr("cx")) * 96 / 914400;
		h = parseInt(ext.attr("cy")) * 96 / 914400;
		console.log([w, h]);
	}
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
							
							// Read relationship file of the slide (Get slideLayoutXX.xml)
							// ppt/slides/slide1.xml
							// ppt/slides/_rels/slide1.xml.rels
							var resName = name.replace("slides/slide", "slides/_rels/slide") + ".rels";
							var $resTarget = openXMLFromZip(zip, resName)
								.find("Relationship[Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\"]")
								.attr("Target")
								.replace("../", "ppt/");
							console.log($resTarget);
							
							// Open slideLayoutXX.xml
							var $slideLayoutXML = openXMLFromZip(zip, $resTarget);
							
							// Parse the slide context and rander into html
							$slideXML.find("sp").each(function(index, element) {
								var $e = $(element);
								var type = $e.find("ph").attr("type");
								var text = $e.find("t").text();
								var id = $e.find("cNvPr").attr("id");
								console.log("  id: " + id);
								
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
									text = "<div class='block content' style='" + getPosition(element, id, $slideLayoutXML) + getSize(element, id, $slideLayoutXML) + "'>";
									$e.find("p").each(function(index, element) {
										text += "<div style='color: " + getFontColor(element) + 
												"; font-size: " + getFontSize(element, id, $slideLayoutXML) + 
												"; font-weight: " + getFontBold(element) + 
												"; font-style: " + getFontItalic(element) + 
												"; font-family: " + getFontType(element) + 
												";'>" + $(element).find("t").text() + "</div>";
									});
									text += "</div>";
								//}
								context += text;
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

			// read the file !
			// readAsArrayBuffer and readAsBinaryString both produce valid content for JSZip.
			reader.readAsArrayBuffer(f);
			// reader.readAsBinaryString(f);
		}
	});
})();
