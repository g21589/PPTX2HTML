$(document).ready(function() {

	if (window.Worker) {
		
		var $result = $("#result");
		var isDone = false;
		
		$("#uploadBtn").on("change", function(evt) {
			
			isDone = false;
			
			$result.html("");
			$("#load-progress").text("0%").attr("aria-valuenow", 0).css("width", "0%");
			$("#result_block").removeClass("hidden").addClass("show");
			
			var fileName = evt.target.files[0];
			
			// Read the file
			var reader = new FileReader();
			reader.onload = (function(theFile) {
				return function(e) {
					
					// Web Worker
					var worker = new Worker('./js/worker.js');
				
					worker.addEventListener('message', function(e) {
						
						var msg = e.data;
						
						switch(msg.type) {
							case "progress-update":
								$("#load-progress").text(msg.data.toFixed(2) + "%")
									.attr("aria-valuenow", msg.data.toFixed(2))
									.css("width", msg.data.toFixed(2) + "%");
								break;
							case "slide":
								$result.append(msg.data);
								break;
							case "processMsgQueue":
								var queue = msg.data;
								for (var i=0; i<queue.length; i++) {
									
									queue[i].data.chartID;
									queue[i].data.chartType;
									var d = queue[i].data.chartData;
									/*
									var sin = [], cos = [];
									
									for (var j = 0; j < 100; j++) {
										sin.push({x: j, y: Math.sin(j/10)});
										cos.push({x: j, y: .5 * Math.cos(j/10)});
									}
									
									var data =  [{
										values: sin,
										key: 'Sine Wave',
										color: '#ff7f0e'
									}, {
										values: cos,
										key: 'Cosine Wave',
										color: '#2ca02c'
									}];
									*/
									var data =  [];
									for (var j=0; j<d.length; j++) {
										var arr = [];
										for (var k=0; k<d[j].length; k++) {
											arr.push({x: k, y: d[j][k]});
										}
										data.push({
											key: 'data' + (j + 1),
											values: arr
										});
									}
									
									var chart = nv.models.lineChart()
										.useInteractiveGuideline(true);
									
									chart.xAxis
										.axisLabel('X')
										.tickFormat(d3.format(',r'));
									
									chart.yAxis
										.axisLabel('Y')
										.tickFormat(d3.format('.02f'));
									
									//document.getElementById("#" + queue[i].data.chartID).innerHTML = "";
									d3.select("#" + queue[i].data.chartID)
										.append("svg")
										.datum(data)
										.transition().duration(500)
										.call(chart);
									
									nv.utils.windowResize(chart.update);
									
								}
								break;
							case "pptx-thumb":
								$("#pptx-thumb").attr("src", "data:image/jpeg;base64," + msg.data);
								break;
							case "ExecutionTime":
								$("#info_block").html("Execution Time: " + msg.data + " (ms)");
								isDone = true;
								worker.postMessage({
									"type": "getMsgQueue"
								});
								break;
							case "WARN":
								console.warn('Worker: ', msg.data);
								break;
							case "ERROR":
								console.error('Worker: ', msg.data);
								$("#error_block").text(msg.data);
								break;
							case "DEBUG":
								console.debug('Worker: ', msg.data);
								break;
							case "INFO":
							default:
								console.info('Worker: ', msg.data);
								//$("#info_block").html($("#info_block").html() + "<br><br>" + msg.data);
						}
						
					}, false);
					
					worker.postMessage({
						"type": "processPPTX",
						"data": e.target.result
					});
					
				}
			})(fileName);
			reader.readAsArrayBuffer(fileName);
			
		});
		
		$("#slideContentModel").on("show.bs.modal", function (e) {
			if (!isDone) { return; }
			$("#slideContentModel .modal-body textarea").text($result.html());
		});
		
		$("#download-btn").click(function () {
			if (!isDone) { return; }
			var cssText = "";
			$.get("css/pptx2html.css", function (data) {
				cssText = data;
			}).done(function () {
				var headHtml = "<style>" + cssText + "</style>";
				var bodyHtml = $result.html();
				var html = "<!DOCTYPE html><html><head>" + headHtml + "</head><body>" + bodyHtml + "</body></html>";
				var blob = new Blob([html], {type: "text/html;charset=utf-8"});
				saveAs(blob, "slides_p.html");
			});
		});
		
		$("#download-reveal-btn").click(function () {
			if (!isDone) { return; }
			var cssText = "";
			$.get("css/pptx2html.css", function (data) {
				cssText = data;
			}).done(function () {
				var revealPrefix = 
"<script type='text/javascript'>\
Reveal.initialize({\
	controls: true,\
	progress: true,\
	history: true,\
	center: true,\
	keyboard: true,\
	slideNumber: true,\
	\
	theme: Reveal.getQueryHash().theme,\
	transition: Reveal.getQueryHash().transition || 'default',\
	\
	dependencies: [\
		{ src: 'lib/js/classList.js', condition: function() { return !document.body.classList; } },\
		{ src: 'plugin/markdown/marked.js', condition: function() { return !!document.querySelector( '[data-markdown]' ); } },\
		{ src: 'plugin/markdown/markdown.js', condition: function() { return !!document.querySelector( '[data-markdown]' ); } },\
		{ src: 'plugin/highlight/highlight.js', async: true, callback: function() { hljs.initHighlightingOnLoad(); } },\
		{ src: 'plugin/zoom-js/zoom.js', async: true, condition: function() { return !!document.body.classList; } },\
		{ src: 'plugin/notes/notes.js', async: true, condition: function() { return !!document.body.classList; } }\
	]\
});\
</script>";
				var headHtml = "<style>" + cssText + "</style>";
				var bodyHtml = "<div id='slides' class='slides'>" + $result.html() + "</div>";
				var html = revealPrefix + headHtml + bodyHtml;
				var blob = new Blob([html], {type: "text/html;charset=utf-8"});
				saveAs(blob, "slides.html");
			});
		});
		
		$("#to-reveal-btn").click(function () {
			if (localStorage) {
				localStorage.setItem("slides", $result.html());
				window.open("./reveal/demo.html", "_blank");
			} else {
				alert("Browser don't support Web Storage!");
			}
		});
		
	} else {
		
		alert("Browser don't support Web Worker!");
		
	}
	
});
