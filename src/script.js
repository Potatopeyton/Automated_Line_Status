    const dropZone = document.getElementById('dropZone');
    const fileList = document.getElementById('fileList');
    const processButton = document.getElementById('processButton');
    let uploadedFiles = [];

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.style.backgroundColor = '#555';
    });

    dropZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        dropZone.style.backgroundColor = '#444';
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.style.backgroundColor = '#444';
        const files = Array.from(e.dataTransfer.files);
        handleFiles(files);
    });

    dropZone.addEventListener('click', () => {
        const fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.accept = '.csv';
        fileInput.onchange = (e) => {
            const files = Array.from(fileInput.files);
            handleFiles(files);
        };
        fileInput.click();
    });

processButton.addEventListener('click', () => {
    var projectedInput = parseFloat(document.getElementById('projectedInput').value);
    if (uploadedFiles.length >= 1) { // Check if at least one file is uploaded
        if (projectedInput > 0) {     // Checks to see if a Pos number is inputed
            processAllFiles();
        } else {
            alert('Please enter a projected value');
            console.log(projectedInput);
        }
    } else {
        alert('Please upload at least one CSV file before processing.');
    }
});
function handleFiles(files) {
        files.forEach(file => {
            if (!file.name.endsWith('.csv')) {
                alert('Only CSV files are allowed');
                return;
            }
            uploadedFiles.push(file);
            const fileItem = document.createElement('li');
            fileItem.className = 'file-item';
            fileItem.textContent = file.name;
            // Apply color coding based on file name
            if (file.name.includes('191')) {
                fileItem.classList.add('red');          // Sets file to red if OP
            } else if (file.name.includes('192')) {
                fileItem.classList.add('dark-yellow');     // Sets file to Dark Yellow if BB
            }

            const removeButton = document.createElement('button');
            removeButton.textContent = 'Remove';
            removeButton.className = 'remove-button';
            removeButton.onclick = function() {
                fileList.removeChild(fileItem);
                uploadedFiles = uploadedFiles.filter(f => f !== file);
            };
            fileItem.appendChild(removeButton);
            fileList.appendChild(fileItem);
        });
    }
function processAllFiles() {
    if (!uploadedFiles.length) {
        alert('Please upload at least one CSV file before processing.');
        return;
    }

    console.log('Processing all files...');
    const wb = XLSX.utils.book_new();
    uploadedFiles.sort((a, b) => a.name.localeCompare(b.name));

    const filePromises = uploadedFiles.map(file => processFile(file, wb));

    Promise.all(filePromises)
        .then(() => {
            createFinalSheet(wb);
        })
        .catch(error => {
            console.error('Error processing files:', error);
            alert('Error processing files: ' + error);
        })
        .finally(() => {
		console.log('Processing completed.');
        });
}

    function processFile(file, wb) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = e.target.result;
                    const parsedCSV = Papa.parse(data, {header: true}).data;
                    let sheetName = file.name.includes("191") ? "OP" : file.name.includes("192") ? "BB" : "Sheet1";
                    const ws = processCSV(parsedCSV);
                    XLSX.utils.book_append_sheet(wb, ws, sheetName);
                    resolve();
                } catch (error) {
                    reject(`Failed to process file ${file.name}: ${error}`);
                }
            };
            reader.onerror = () => reject(`Error reading file ${file.name}`);
            reader.readAsText(file);
        });
    }

    function processCSV(data) {
        let hourData = {};
        data.forEach(row => {
            let time = row['@timestamp per 30 minutes'];
            let count = row['Count'] ? parseInt(row['Count'].replace(',', ''), 10) : 0;
            if (time && !isNaN(count)) {
                let [hour, minute] = time.split(':');
            let hourKey = `${hour}:00-${hour}:59`;
                hourData[hourKey] = hourData[hourKey] || { firstHalf: 0, secondHalf: 0 };
                hourData[hourKey][minute < 30 ? 'firstHalf' : 'secondHalf'] += count;
            }
        });

        const processedData = Object.entries(hourData).map(([key, {firstHalf, secondHalf}]) => ({
            Time: key,
            '1st half': firstHalf,
            '2nd half': secondHalf,
            'Total': firstHalf + secondHalf
        }));
        return XLSX.utils.json_to_sheet(processedData);
    }
function createFinalSheet(wb) {
    try {
        const projectedValue = parseFloat(document.getElementById('projectedInput').value) || 0;
        let bbData = wb.Sheets['BB'] ? XLSX.utils.sheet_to_json(wb.Sheets['BB'], { header: 1 }).slice(1) : [];
        let opData = wb.Sheets['OP'] ? XLSX.utils.sheet_to_json(wb.Sheets['OP'], { header: 1 }).slice(1) : [];
        let timeSlots = new Set([...bbData.map(row => row[0]), ...opData.map(row => row[0])]);

        // Custom sorting logic to wrap around after midnight starting from 07:00 to 04:00
        let sortedTimeSlots = Array.from(timeSlots).sort((a, b) => {
            let hourA = parseInt(a.split(':')[0], 10);
            let hourB = parseInt(b.split(':')[0], 10);
            // Adjust hours for sorting: shift times after 04:00 to the previous logical day
            if (hourA >= 7) hourA -= 24;
            if (hourB >= 7) hourB -= 24;
            return hourA - hourB;
        });

        const finalData = [['Time', 'BB 1st half', 'BB 2nd half', 'BB Total', 'OP 1st half', 'OP 2nd half', 'OP Total', 'Hourly Total', 'Running Total', 'Projected Remaining', 'Projected']];
        let runningTotal = 0;

        const bbMap = new Map(bbData.map(row => [row[0], row]));
        const opMap = new Map(opData.map(row => [row[0], row]));

        sortedTimeSlots.forEach(time => {
            let bbRow = bbMap.get(time) || [time, 0, 0, 0];
            let opRow = opMap.get(time) || [time, 0, 0, 0];
            let hourlyTotal = (bbRow[3] || 0) + (opRow[3] || 0);
            runningTotal += hourlyTotal;
            let difference = projectedValue - runningTotal;

            finalData.push([
                time,
                (bbRow[1] || 0).toLocaleString(),
                (bbRow[2] || 0).toLocaleString(),
                (bbRow[3] || 0).toLocaleString(),
                (opRow[1] || 0).toLocaleString(),
                (opRow[2] || 0).toLocaleString(),
                (opRow[3] || 0).toLocaleString(),
                hourlyTotal.toLocaleString(),
                runningTotal.toLocaleString(),
                difference.toLocaleString(),
                finalData.length === 1 ? projectedValue.toLocaleString() : ""
            ]);
        });

        const finalSheet = XLSX.utils.aoa_to_sheet(finalData);
        setStyle(finalSheet); // Apply styles to the final sheet
        XLSX.utils.book_append_sheet(wb, finalSheet, 'Final');
        saveWorkbook(wb);
	console.log('Final', finalSheet);
	console.log('BB', bbData);
	console.log('OP', opData);
    } catch (error) {
        console.error('An error occurred in createFinalSheet:', error);
        alert('An error occurred while processing the final sheet: ' + error.message);
    }
}

function setStyle(ws) {
    const borderStyle = {
        top: { style: 'thin', color: { rgb: "000000" } },
        bottom: { style: 'thin', color: { rgb: "000000" } },
        left: { style: 'thin', color: { rgb: "000000" } },
        right: { style: 'thin', color: { rgb: "000000" } }
    };

    const headerAndTimeCellStyle = {
        font: {
            name: "Calibri",
            sz: 8,
            bold: true,
            color: { rgb: "000000" }
        },
        fill: {
            type: "pattern",
            pattern: "solid",
            fgColor: { rgb: "ADD8E6" }
        },
        alignment: {
            horizontal: "center",
            vertical: "center",
            wrapText: true
        },
        border: borderStyle
    };

    // Define column widths and hide columns
    ws['!cols'] = ws['!cols'] || [];
    ws['!cols'][0] = { width: 12 }; // Set the width of column A to 12
    ws['!cols'][1] = { hidden: true }; // Hide column B
    ws['!cols'][2] = { hidden: true }; // Hide column C
    ws['!cols'][4] = { hidden: true }; // Hide column E
    ws['!cols'][5] = { hidden: true }; // Hide column F

    // Iterate through all cells to apply styles
    for (let row in ws) {
        if (row === "!ref") continue;
        let cell = ws[row];
        let rowNum = parseInt(row.replace(/[A-Z]/g, ''), 10); // Extract the row number

        if (cell.t) { // If the cell has a type
            // Apply header and time cell styles
            if (row.match(/^[A-Z]1$/)) { // Assuming headers are in the first row
                cell.s = headerAndTimeCellStyle;
            } else if (row.match(/^[A]\d+$/)) { // Apply style to all rows in column A
                cell.s = headerAndTimeCellStyle;
            } else {
                // General style for all other cells
                cell.s = {
                    font: {
                        name: "Calibri",
                        sz: 8,
                        bold: false,
                        color: { rgb: "000000" }
                    },
                    alignment: {
                        horizontal: "center",
                        vertical: "center",
                        wrapText: true
                    },
                    border: borderStyle
                };
            }
            if (row === 'D1') {
                cell.s = {
                    ...headerAndTimeCellStyle,
                    fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FFDE2A" } } // dark yellow
                };
            }
            if (row === 'G1') {
                cell.s = {
                    ...headerAndTimeCellStyle,
                    fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FF5B5B" } } // opaque red
                };
            }
            // Specific styling for cell 'K2'
            if (row === 'K2') {
                cell.s = { font: {
                        name: "Calibri",
                        sz: 8,
                        bold: false,
                        color: { rgb: "000000" }
                    },
                    alignment: {
                        horizontal: "center",
                        vertical: "center",
                        wrapText: true
                    },
                    border: borderStyle
                }; // Apply cell style to K2
            }

            // No styling for cells 'K3' onwards in the 'Projected' column
            if (row.startsWith('K') && rowNum > 2) {
                cell.s = {}; // Apply no styles
            }
        }
    }
}

    function saveWorkbook(wb) {
        const now = new Date();
	 if (now.getHours() < 5) {
        now.setDate(now.getDate() - 1); // Set to the previous day because files generated from 00:00- 03:59 are for previous day
    	}
        const year = now.getFullYear();
        const month = now.toLocaleString('default', { month: 'short' }).toUpperCase();
        const date = now.getDate();
        const hour = now.getHours();
        const filename = `${date} ${month} ${year} ${hour}:00.xlsx`;
        XLSX.writeFile(wb, filename);
    }
		                                                                                                                //Peyton Lyszczarz peyton.lyszczarz@davita.com
