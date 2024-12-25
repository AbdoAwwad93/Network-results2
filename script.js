let examResults = {};

window.addEventListener('load', function() {
    fetch('./Net- midterm grads 2024-2025.xlsx')
        .then(response => response.arrayBuffer())
        .then(data => {
            var workbook = XLSX.read(data, {type: 'array'});
            var firstSheetName = workbook.SheetNames[0];
            var worksheet = workbook.Sheets[firstSheetName];
            var jsonData = XLSX.utils.sheet_to_json(worksheet);

            examResults = {};
            jsonData.forEach(row => {
                var studentNumber = row['رقم الطالب'];
                var result = row['Total MidTerm Grad'];
                examResults[studentNumber] = result;
            });

            console.log('Excel file processed:', examResults);
        })
        .catch(error => console.error('Error loading Excel file:', error));
});

function lookupResult() {
    const studentNumber = document.getElementById('studentNumber').value;
    const resultDiv = document.getElementById('result');
    
    if (examResults.hasOwnProperty(studentNumber)) {
        resultDiv.textContent = `Result: ${examResults[studentNumber]}`;
    } else {
        resultDiv.textContent = "No result found for this student number.";
    }
} 