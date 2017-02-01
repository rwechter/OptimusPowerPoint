// The initialize function is required for all add-ins.
Office.initialize = function (reason) {

    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {

        // Execute sendFile when submit is clicked
        $('#submit').click(function () {
            sendFile();
        });

        // Update status
        updateStatus("Ready to send file.");
    });
}

// Create a function for writing to the status div.
function updateStatus(message) {
    var statusInfo = $('#status');
    statusInfo.innerHTML += message + "<br/>";
}
        function writeText(event) {

            //Consult Office.js API reference to see all you can do. This just shows the simplest action. 

            Office.context.document.setSelectedDataAsync("ExecuteFunction Works. Button ID=" + event.source.id,
                function (asyncResult) {
                    var error = asyncResult.error;
                    if (asyncResult.status === "failed") {
                        //show error. Upcoming displayDialog API will help here.
                    }
                    else {
                        //show success.Upcoming displayDialog API will help here.
                    }
                });
           //Required, call event.completed to let the platform know you are done processing. 
	   event.completed();
        }
