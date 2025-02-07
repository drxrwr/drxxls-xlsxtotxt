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
        outputDiv.innerHTML = ""; // Bersihkan output sebelumnya

        workbook.SheetNames.forEach((sheetName) => {
            let sheet = workbook.Sheets[sheetName];
            let rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // Ambil data sebagai array

            if (rows.length === 0) {
                alert("Sheet kosong!");
                return;
            }

            // Transpose data untuk mendapatkan kolom per kolom
            let numColumns = rows[0].length;
            for (let col = 0; col < numColumns; col++) {
                let columnData = rows.map(row => row[col]).filter(cell => cell !== undefined).join("\n");

                let blob = new Blob([columnData], { type: 'text/plain' });
                let url = URL.createObjectURL(blob);
                
                let columnName = String.fromCharCode(65 + col); // A, B, C, ...
                let downloadLink = document.createElement("a");
                downloadLink.href = url;
                downloadLink.download = file.name.replace(/\.[^/.]+$/, "") + `_${columnName}.txt`;
                downloadLink.innerText = `Download TXT - ${columnName}`;
                downloadLink.style.display = "block";
                outputDiv.appendChild(downloadLink);
            }
        });
    };

    reader.readAsArrayBuffer(file);
}
