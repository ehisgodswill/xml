const fs = require('fs');
const XlsxPopulate = require('xlsx-populate');
// Function to read JSON file
function readJsonFile(filePath) {
    try {
        const data = fs.readFileSync(filePath, 'utf8');
        return JSON.parse(data);
    } catch (err) {
        console.error('Error reading JSON file:', err);
        return {};
    }
}

// Function to write JSON file
function processFiles (jsonData) {
    XlsxPopulate.fromBlankAsync().then((workbook) => {
        const sheet = workbook.sheet(0);

        sheet.column("A").width(8).cell(1).value("FarmerID");
        sheet.column("B").width(10).cell(1).value("Status");
        sheet.column("C").width(30).cell(1).value("Name");
        sheet.column("D").width(8).cell(1).value("Sex");
        sheet.column("E").width(10).cell(1).value("Farm");
        sheet
            .column("F")
            .width(14)
            .style({ horizontalAlignment: "justify" })
            .cell(1)
            .value("Phone No");
        sheet
            .column("G")
            .width(12)
            .hidden(true)
            .cell(1)
            .value("");
        sheet
            .column("H")
            .width(25)
            .hidden(true)
            .style({ wrapText: false })
            .cell(1)
            .value("Cordinates");
        sheet.column("I").width(9).cell(1).value("Hectrage");
        sheet.column("J").width(15).cell(1).value("Percentage");
        sheet.column("K").width(9).cell(1).value("Area (Ha)");
        sheet
            .column("L")
            .width(16)
            .style({ horizontalAlignment: "center", wrapText: true })
            .cell(1)
            .value("Crop");
        sheet.column("M").width(50).cell(1).value("Farmer Photo");
        sheet.column("N").width(50).cell(1).value("Farm Photo");
        sheet.column("O").hidden(true).width(10).cell(1).value("uuid");
        sheet.column("P").width(25).cell(1).value("Date");
        sheet.row(1).style({ fill: 'cccccc', bold: true, fontSize: 14 });
        sheet.freezePanes(0, 1);

        let i = 0;
        for (feature of jsonData.features) {
            fillSheet(sheet, feature.properties, i++);
        }

        workbook.outputAsync({ type: "blob" }).then((blob) => {
            if (window.navigator && window.navigator.msSaveOrOpenBlob) {
                window.navigator.msSaveOrOpenBlob(blob, "Data.xlsx");
            } else {
                var url = window.URL.createObjectURL(blob);
                var a = document.createElement("a");
                document.body.appendChild(a);
                a.href = url;
                a.download = "Data.xlsx";
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
            }
        });
    });
}

// Function to fill sheet with data
function fillSheet (sheet, obj, index) {
    const maxrowNum = sheet.usedRange()
        ? sheet.usedRange()._maxRowNumber + 1
        : 1;
    const row = sheet.row(maxrowNum);
    const range = sheet.range(maxrowNum, "A", maxrowNum, "P");

    range.forEach((cell) => {
        switch (cell.columnName()) {
            case "A":
                const date = new Date(obj.start);
                cell.value(
                    `0${obj.Collector_Id}${String(obj.Farmers_ID).length < 2
                        ? "0" + obj.Farmers_ID
                        : obj.Farmers_ID
                    }-${date.getDate()}/${date.getMonth() + 1}/${date.getFullYear() - 2000}`
                );
                break;
            case "B":
                cell.value(obj.Respondent_Status);
                break;
            case "C":
                cell.value(
                    `${obj.First_Name.trim()} ${obj.Last_Name.trim()} ${!obj.Other_Names || obj.Other_Names === "-"
                        ? ""
                        : obj.Other_Names.trim()
                        }`.trim()
                );
                break;
            case "D":
                cell.value(obj.Sex);
                break;
            case "E":
                cell.value("Farm" + obj.FarmStatus);
                break;
            case "F":
                cell.value(String(obj.Phone_No || '').trim());
                break;
            case "G":
                // cell.value(
                //   `0${obj.Collector_Id}${String(obj.Farmers_ID).length < 2
                //     ? "0" + obj.Farmers_ID
                //     : obj.Farmers_ID
                //   }`);
                break;
            case "H":
                const cords = obj.FarmBoundary.trim().split(";").map((cord) => cord.trim().split(" ").slice(0, 2).join(" "));
                cell.value(`POLYGON (( ${cords.join(", ")} ))`);
                break;
            case "I":
                cell.value(Number(obj.Total_Area_Ha));
                break;
            case "J":
                cell.value(obj.AreaCropped + '%');
                break;
            case "K":
                cell.value(Number(parseFloat(obj.Total_Area_Ha * obj.AreaCropped * 0.01).toFixed(2)));
                break;
            case "L":
                // resultObj.crop[obj._index].forEach((crop, i) => {
                //   sheet.cell(maxrowNum + i, "L").value(crop);
                // });
                cell.value(resultObj.crop[obj._index].join("\n"));
                break;
            case "M":
                cell.value(obj.Farmers_Picture);
                break;
            case "N":
                cell.value(obj.Density_Picture);
                break;
            case "O":
                cell.value(obj._uuid);
                break;
            case "P":
                cell.value(new Date(obj.start).toDateString());
                break;
            default:
                break;
        }
    });
    sheet
        .range(maxrowNum, "M", sheet.usedRange()._maxRowNumber, "M")
        .merged(true);
    sheet
        .range(maxrowNum, "N", sheet.usedRange()._maxRowNumber, "N")
        .merged(true);
    sheet
        .row(sheet.usedRange()._maxRowNumber).height(150)
        .style({ bottomBorder: { style: "thin", color: "0f0f0f" } });
    row.style({ fill: index % 2 ? 'f0f0f0' : 'ffffff' });

    for (const key in sheet._dataValidations) {
        if (key[1] === "1") {
            const newKey = key + " " + key + maxrowNum;
            const value = sheet._dataValidations[key];

            value.attributes.sqref = newKey;
            sheet._dataValidations[newKey] = value;
            delete sheet._dataValidations[key];
        }
    }
}

// Function to process JSON data
function processJsonData(jsonData) {
    // Implement your processing logic here

}

// Example usage
const filePath = 'C:\\Users\\USER\\Downloads\\iguere.geojson';
const jsonData = readJsonFile(filePath);
if (jsonData) {
    processJsonData(jsonData);
}




