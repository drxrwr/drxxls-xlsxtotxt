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
            let rows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });

            if (rows.length === 0) {
                alert("Sheet kosong!");
                return;
            }

            // Cari jumlah kolom maksimum
            let maxCols = Math.max(...rows.map(row => row.length));

            // Buat array untuk menyimpan data per kolom
            let columnDataArray = Array.from({ length: maxCols }, () => []);

            // Masukkan data ke array kolom secara berurutan
            rows.forEach(row => {
                for (let col = 0; col < maxCols; col++) {
                    columnDataArray[col].push(row[col] || ""); // Gunakan "" jika kosong
                }
            });

            // Simpan setiap kolom sebagai file TXT
            columnDataArray.forEach((columnData, colIndex) => {
                let txtData = columnData.join("\n");

                let blob = new Blob([txtData], { type: 'text/plain' });
                let url = URL.createObjectURL(blob);

                let columnName = String.fromCharCode(65 + colIndex); // A, B, C, ...
                let downloadLink = document.createElement("a");
                downloadLink.href = url;
                downloadLink.download = file.name.replace(/\.[^/.]+$/, "") + `_${columnName}.txt`;
                downloadLink.innerText = `Download TXT - ${columnName}`;
                downloadLink.style.display = "block";
                outputDiv.appendChild(downloadLink);
            });
        });
    };

    reader.readAsArrayBuffer(file);
}
