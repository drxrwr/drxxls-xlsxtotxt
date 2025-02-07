function convertExcelToTxt() {
    let input = document.getElementById('fileInput');
    if (input.files.length === 0) {
        alert('Pilih file Excel terlebih dahulu!');
        return;
    }

    let file = input.files[0];
    let reader = new FileReader();

    reader.onload = function (e) {
        let data = new Uint8Array(e.target.result);
        let workbook = XLSX.read(data, { type: 'array' });

        let outputDiv = document.getElementById('output');
        outputDiv.innerHTML = ""; // Kosongkan output sebelumnya

        workbook.SheetNames.forEach((sheetName) => {
            let sheet = workbook.Sheets[sheetName];
            let txtData = XLSX.utils.sheet_to_csv(sheet, { FS: '\t' }); // Gunakan tab sebagai pemisah

            let blob = new Blob([txtData], { type: 'text/plain' });
            let url = URL.createObjectURL(blob);
            
            let downloadLink = document.createElement("a");
            downloadLink.href = url;
            downloadLink.download = file.name.replace(/\.[^/.]+$/, "") + "_" + sheetName + ".txt";
            downloadLink.innerText = "Download TXT - " + sheetName;
            downloadLink.style.display = "block";
            outputDiv.appendChild(downloadLink);
        });
    };

    reader.readAsArrayBuffer(file);
}
