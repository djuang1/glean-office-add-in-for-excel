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
    const field1Value = document.getElementById('field1').value;
    const field2Value = document.getElementById('field2').value;

    if (!field1Value.trim() && !field2Value.trim()) {
        showStatus('Please enter at least one value', 'error');
        return;
    }

    window.gleanGlobals.instance = field1Value;
    window.gleanGlobals.token = field2Value;

    showStatus(`Configuration applied`, 'success');
    
    // try {
    //     await Excel.run(async (context) => {
    //         // Get the active worksheet
    //         const worksheet = context.workbook.worksheets.getActiveWorksheet();

    //         // Get the selected range or use A1:B1 as default
    //         const range = context.workbook.getSelectedRange();
    //         range.load("address");

    //         await context.sync();

    //         // Insert data into the selected range
    //         const values = [[field1Value, field2Value]];
    //         range.values = values;

    //         // Auto-fit columns
    //         range.format.autofitColumns();

    //         await context.sync();

    //         showStatus(`Data inserted successfully at ${range.address}`, 'success');
    //     });
    // } catch (error) {
    //     console.error('Error inserting data:', error);
    //     showStatus('Error inserting data: ' + error.message, 'error');
    // }
}

function clearFields() {
    document.getElementById('field1').value = '';
    document.getElementById('field2').value = '';
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
document.getElementById('field1').addEventListener('keypress', function (e) {
    if (e.key === 'Enter') {
        document.getElementById('field2').focus();
    }
});

document.getElementById('field2').addEventListener('keypress', function (e) {
    if (e.key === 'Enter') {
        insertData();
    }
});