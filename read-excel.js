function readExcelFile(file, callback) {
    var reader = new FileReader();
    reader.onload = function(e) {
        var data = e.target.result;
        var workbook = XLSX.read(data, {
            type: 'binary'
        });
        var first_sheet_name = workbook.SheetNames[0];
        var worksheet = workbook.Sheets[first_sheet_name];
        var result = XLSX.utils.sheet_to_json(worksheet);
        callback(result);
    };
    reader.readAsBinaryString(file);
}

