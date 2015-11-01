importScripts(
	'./js/jszip.min.js',
	'./js/highlight.min.js',
	'./js/colz.class.min.js',
	'./js/highlight.min.js',
	'./js/functions.js',
	'./js/tXmlUnfolded.min.js'
);

onmessage = function(e) {
	
	var zip = new JSZip(e.data);
	
	self.postMessage( base64ArrayBuffer(zip.file("docProps/thumbnail.jpeg").asArrayBuffer() ));
	
	var xmlObj = tXml(zip.file("ppt/presentation.xml").asText());
	
	self.postMessage( xmlObj );
	
}
