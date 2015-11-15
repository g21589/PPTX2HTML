$(document).ready(function() {

	if (window.Worker) {
		
		var $result = $("#result");
		
		$("#uploadBtn").on("change", function(evt) {
			
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
							case "pptx-thumb":
								$("#pptx-thumb").attr("src", "data:image/jpeg;base64," + msg.data);
								break;
							case "ExecutionTime":
								$("#info_block").html("Execution Time: " + msg.data + " (ms)");
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
					
					worker.postMessage(e.target.result);
					
				}
			})(fileName);
			reader.readAsArrayBuffer(fileName);
			
		});
		
		$("#slideContentModel").on("show.bs.modal", function (e) {
			$("#slideContentModel .modal-body textarea").text($result.html());
		});
		
	} else {
		
		alert("Browser don't support Web Worker!");
		
	}
	
});
