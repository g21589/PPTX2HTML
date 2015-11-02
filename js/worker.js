importScripts(
	'./jszip.min.js',
	'./highlight.min.js',
	'./colz.class.min.js',
	'./highlight.min.js',
	'./functions.js',
	'./tXmlUnfolded.min.js'
);

onmessage = function(e) {
	
	var zip = new JSZip(e.data);
	
	self.postMessage({
		"type": "pptx-thumb",
		"data": base64ArrayBuffer(zip.file("docProps/thumbnail.jpeg").asArrayBuffer())
	});
	
	var jsonContent = tXml(zip.file("ppt/presentation.xml").asText(), { simplify: 1 });
	//var jsonContent = xmlToJSON.parseString( "", {xmlns: false}/*zip.file("ppt/presentation.xml").asText()*/ );
	
	self.postMessage({
		"type": "INFO",
		"data": JSON.stringify(jsonContent)
	});
	
}
