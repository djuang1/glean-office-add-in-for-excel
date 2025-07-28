Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log("Glean Excel Add-in loaded successfully");

        if (typeof CustomFunctions !== "undefined") {
            console.log("Custom functions registered");
        }

        OfficeRuntime.storage.getItem("instance").then(function (result) {
            console.log("Success: Item with key '" + "instance" + "' read from Storage.");
            document.getElementById('instanceName').value = result;
            window.gleanGlobals.instance = result;
        }, function (error) {
            console.log("Error: Unable to read item with key '" + "instance" + "' from Storage. " + error);
        });

        OfficeRuntime.storage.getItem("token").then(function (result) {
            console.log("Success: Item with key '" + "token" + "' read from Storage.");
            document.getElementById('apiToken').value = result;
            window.gleanGlobals.token = result;
        }, function (error) {
            console.log("Error: Unable to read item with key '" + "token" + "' from Storage. " + error);
        });
    }
});

window.gleanGlobals = {
    instance: "foo",
    token: "bar"
};

async function insertData() {
    const instanceNameValue = document.getElementById('instanceName').value;
    const apiTokenValue = document.getElementById('apiToken').value;

    if (!instanceNameValue.trim() && !apiTokenValue.trim()) {
        showStatus('Please enter at least one value', 'error');
        return;
    }

    OfficeRuntime.storage.setItem("instance", instanceNameValue).then(function () {
        console.log("Success: Item with key 'instance' saved to Storage.");
    }, function (error) {
        console.log("Error: Unable to save item with key '" + "instance" + "' to Storage. " + error);
    });

    OfficeRuntime.storage.setItem("token", apiTokenValue).then(function () {
        console.log("Success: Item with key '" + "token" + "' saved to Storage.");
        showStatus(`Configuration saved`, 'success');
    }, function (error) {
        console.log("Error: Unable to save item with key '" + "token" + "' to Storage. " + error);
        showStatus(`Error saving configuration: ${error}`, 'error');
    });

    window.gleanGlobals.instance = instanceNameValue;
    window.gleanGlobals.token = apiTokenValue;

    showStatus(`Configuration applied`, 'success');
}

function clearFields() {
    document.getElementById('instanceName').value = '';
    document.getElementById('apiToken').value = '';
    hideStatus();
}

function showStatus(message, type) {
    const statusDiv = document.getElementById('status');
    statusDiv.textContent = message;
    statusDiv.className = `status ${type}`;
    statusDiv.style.display = 'block';

    // Auto-hide success messages after 3 seconds
    if (type === 'success') {
        setTimeout(hideStatus, 3000);
    }
}

function hideStatus() {
    const statusDiv = document.getElementById('status');
    statusDiv.style.display = 'none';
}

// Handle Enter key in input fields
document.getElementById('instanceName').addEventListener('keypress', function (e) {
    if (e.key === 'Enter') {
        document.getElementById('apiToken').focus();
    }
});

document.getElementById('apiToken').addEventListener('keypress', function (e) {
    if (e.key === 'Enter') {
        insertData();
    }
});