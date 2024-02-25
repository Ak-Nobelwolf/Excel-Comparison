document.getElementById('compareBtn').addEventListener('click', function(event) {
    // Check if any files have been uploaded
    var sourceFile = document.getElementById('source').value;
    var targetFile = document.getElementById('target').value;
    var configFile = document.getElementById('config').value;

    if (!sourceFile && !targetFile && !configFile) {
        alert('Please upload all three files before comparing.');
        event.preventDefault();
        return;
    }

    // File Type Validation for Source File
    if (sourceFile) {
        if (!(sourceFile.endsWith('.xlsx') || sourceFile.endsWith('.xls'))) {
            alert('Invalid file type for the Source File. Please upload an Excel file.');
            event.preventDefault();
            return;
        }
    } else {
        alert('Please upload a Source File.');
        event.preventDefault();
        return;
    }

    // File Type Validation for Target File
    if (targetFile) {
        if (!(targetFile.endsWith('.xlsx') || targetFile.endsWith('.xls'))) {
            alert('Invalid file type for the Target File. Please upload an Excel file.');
            event.preventDefault();
            return;
        }
    } else {
        alert('Please upload a Target File.');
        event.preventDefault();
        return;
    }

    // File Type Validation for Config File
    if (configFile) {
        if (!configFile.endsWith('.properties')) {
            alert('Invalid file type for the Config File. Please upload a properties file.');
            event.preventDefault();
            return;
        }
    } else {
        alert('Please upload a Config File.');
        event.preventDefault();
        return;
    }
});

// Function to hide loader and display alert message
function hideLoaderAndDisplayMessage() {
  // Hide loader
  document.getElementById('loader').style.display = 'none';
  // Display alert message
  alert('Comparison completed. Check the result.');
  // Reload the page
  location.reload();
}