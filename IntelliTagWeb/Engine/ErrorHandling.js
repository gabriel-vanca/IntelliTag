//$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
function errorHandler(error) {
    // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
    showNotification("Error:", error);
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}