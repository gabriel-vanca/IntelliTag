// The initialize function must be run each time a new page is loaded.
(function() {
    Office.initialize = function(reason) {
        // If you need to initialize something you can do so here.
    };


})();


//Notice function needs to be in global namespace

function SetUnsetDeontic_OnClick(event) {
    setLogic(setDeonticMarker);
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}
function SetUnsetTemporal_OnClick(event) {
    setLogic(setTemporalMarker);
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}
function SetUnsetOperational_OnClick(event) {
    setLogic(setOperationalMarker);
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}