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
	
	self.postMessage({
		"type": "pptx-thumb",
		"data": base64ArrayBuffer(zip.file("docProps/thumbnail.jpeg").asArrayBuffer())
	});
	
	var ContentTypesJson = readXmlFile(zip, "[Content_Types].xml");
	var subObj = ContentTypesJson["?xml"]["Types"]["Override"];
	for (var i=0; i<subObj.length; i++) {
		switch (subObj[i]["attrs"]["ContentType"]) {
			case "application/vnd.openxmlformats-officedocument.presentationml.slide+xml":
				self.postMessage({
					"type": "INFO",
					"data": subObj[i]["attrs"]["PartName"]
				});
				break;
			default:
		}
		
	}
	//self.postMessage({
	//	"type": "INFO",
	//	"data": JSON.stringify( ContentTypes )
	//});
	
	var presentationJson = readXmlFile(zip, "ppt/presentation.xml");
	//self.postMessage({
	//	"type": "INFO",
	//	"data": JSON.stringify( presentationJson )
	//});
	
}

function readXmlFile(zip, filename) {
	return tXml(zip.file(filename).asText());
}
