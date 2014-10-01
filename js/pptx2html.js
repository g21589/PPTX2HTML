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

function getFontSize(element) {
	var fontSize = (parseInt($(element).find("rPr").attr("sz")) / 100);
	return isNaN(fontSize) ? "inherit" : (fontSize + "pt");
}

function getFontBold(element) {
	return $(element).find("rPr").attr("b") === "1" ? "bold" : "initial";
}

function getFontItalic(element) {
	return $(element).find("rPr").attr("i") === "1" ? "italic" : "normal";
}

function getPosition(element) {
	var off = $(element).find("off");
	var x = parseInt(off.attr("x")) / 10000;
	var y = parseInt(off.attr("y")) / 10000;
	return (isNaN(x) || isNaN(y)) ? "" : "top:" + y + "px; left:" + x + "px;";
}

function getSize(element) {
	var ext = $(element).find("ext");
	var w = parseInt(ext.attr("cx")) / 10000;
	var h = parseInt(ext.attr("cy")) / 10000;
	return (isNaN(w) || isNaN(h)) ? "" : "width:" + w + "px; height:" + h + "px;";
}

function getSlideSize(zip) {
	var $presentationXML = $($.parseXML(zip.file("ppt/presentation.xml").asText()));
	var sizeNode = $presentationXML.find("sldSz");
	return {
		"width": (parseInt(sizeNode.attr("cx")) / 10000),
		"height": (parseInt(sizeNode.attr("cy")) / 10000)
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
						
						// Size information
						var slideSize = getSlideSize(zip);
						
						var slides = zip.file(/slide\d+.xml$/);
						slides.sort(function(a, b) {return parseInt(a.name.substring(16)) - parseInt(b.name.substring(16))});
						
						// that, or a good ol' for(var entryName in zip.files)
						$.each(slides, function (index, zipEntry) {
							
							var context = "";
							
							var xmlDoc = $.parseXML(zipEntry.asText());
							var $xml = $(xmlDoc);
							
							$xml.find("sp").each(function(index, element) {
								var $e = $(element);
								var type = $e.find("ph").attr("type");
								var text = $e.find("t").text();
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
									text = "<div class='block content' style='" + getPosition(element) + getSize(element) + "'>";
									$e.find("p").each(function(index, element) {
										text += "<p style='color: " + getFontColor(element) + 
												"; font-size: " + getFontSize(element) + 
												"; font-weight: " + getFontBold(element) + 
												"; font-style: " + getFontItalic(element) + 
												"; font-family: " + getFontType(element) + 
												";'>" + $(element).find("t").text() + "</p>";
									});
									text += "</div>";
								}
								context += text;
							});
							
							$fileContent.append($("<li>", {
								"class" : "slide",
								html : zipEntry.name + 
									   "<section style='width:" + slideSize.width + "px; height:" + slideSize.height + "px;'>" + context + "</section>" +
									   "<details><pre><code class='xml'>" + escapeHtml(vkbeautify.xml(zipEntry.asText(), 4)) + "</code></pre></details>"
							}));
							
						});
						// end of the magic !
						
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
