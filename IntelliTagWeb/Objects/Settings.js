var Settings = {
    _lastLogicId: 0,

    get lastLogicId() {

        this._lastLogicId = Office.context.document.settings.get('lastLogicId');
        if (this._lastLogicId == null) {
            this._lastLogicId = 0;
            Office.context.document.settings.set('lastLogicId', 0);
            SaveSettings();
        };

        return this._lastLogicId;
    },

    set lastLogicId(id) {
        this._lastLogicId = id;
        Office.context.document.settings.set('lastLogicId', id);
        SaveSettings();
    }
}

//function LoadSettings() {
//        Settings._lastLogicId = Office.context.document.settings.get('lastLogicId');
//}

function SaveSettings() {
    Office.context.document.settings.saveAsync(function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            showNotification("Error while saving",
                "The following error has occured while saving the add-in settings: " + asyncResult.error.message);
        } else {
        }
    });
}