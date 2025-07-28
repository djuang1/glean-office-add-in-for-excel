Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log("Glean Excel Add-in loaded successfully");

        if (typeof CustomFunctions !== "undefined") {
            console.log("Custom functions registered");
        }
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

    const key = "token";
    OfficeRuntime.storage.setItem(key, token).then(function () {
        tokenSendStatus.value = "Success: Item with key '" + key + "' saved to Storage.";
    }, function (error) {
        tokenSendStatus.value = "Error: Unable to save item with key '" + key + "' to Storage. " + error;
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