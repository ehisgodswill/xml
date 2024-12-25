"use strict";
const XlsxPopulate = require("xlsx-populate");
const fs = require("fs");
const excelToJson = require("convert-excel-to-json");

const result = excelToJson({
  sourceFile: "Enumeration.xlsx",
  header: {
    rows: 1,
  },
  sheets: [
    {
      name: "Enumeration",
      columnToKey: Object.fromEntries(
        "abcdefghijklmnopqrstuvwxyz"
          .toUpperCase()
          .split("")
          .map((char) => [char, `{{${char}1}}`])
      ),
    },
    {
      name: "planted_Crops",
      columnToKey: Object.fromEntries(
        "abcdefg"
          .toUpperCase()
          .split("")
          .map((char) => [char, `{{${char}1}}`])
      ),
    },
  ],
});
let planted_Crops = {};
result.planted_Crops.forEach((crop) => {
  const string = `${crop.Assessed_CropPlanted} ${crop.CropPercentage}${
    Number(crop.CropPercentage) ? "%" : ""
  } ${crop.Category_CropPlanted.toUpperCase()};`;

  planted_Crops[crop._parent_index] = planted_Crops[crop._parent_index]
    ? [...planted_Crops[crop._parent_index], string]
    : [string];
});
XlsxPopulate.fromBlankAsync().then((workbook) => {
  const sheet = workbook.sheet(0);

  sheet.column("A").width(8).cell(1).value("FarmerID");
  sheet.column("B").width(10).cell(1).value("Status");
  sheet.column("C").width(30).cell(1).value("Name");
  sheet.column("D").width(8).cell(1).value("Sex");
  sheet.column("E").width(10).cell(1).value("Farm");
  sheet
    .column("F")
    .width(12)
    .style({ horizontalAlignment: "justify" })
    .cell(1)
    .value("Phone No");
  sheet
    .column("G")
    .width(10)
    .style({ horizontalAlignment: "justify" })
    .cell(1)
    .value("X-Cord");
  sheet
    .column("H")
    .width(10)
    .style({ horizontalAlignment: "justify" })
    .cell(1)
    .value("Y-Cord");
  sheet.column("I").width(9).cell(1).value("Area(Ha)");
  sheet.column("J").width(15).cell(1).value("Percentage Cropped");
  sheet
    .column("K")
    .width(16)
    .style({ horizontalAlignment: "center", wrapText: true })
    .cell(1)
    .value("Crop");
  sheet.column("L").width(12).cell(1).value("Community");
  sheet.column("M").width(25).cell(1).value("Farmer Photo");
  sheet.column("N").width(25).cell(1).value("Farm Photo");
  sheet.column("O").hidden(true).width(10).cell(1).value("uuid");
  sheet.column("P").width(25).cell(1).value("Date");

  result.Enumeration.sort((a, b) => {
    return a.Collector_Id < b.Collector_Id
      ? -1
      : a.Collector_Id === b.Collector_Id
      ? a.start < b.start
        ? -1
        : 1
      : 1;
  }).forEach((obj) => {
    const maxrowNum = sheet.usedRange()
      ? sheet.usedRange()._maxRowNumber + 1
      : 1;
    const row = sheet.row(maxrowNum);
    const range = sheet.range(maxrowNum, "A", maxrowNum, "P");

    range.forEach((cell) => {
      switch (cell.columnName()) {
        case "A":
          // console.log(obj);
          cell.value(
            `0${obj.Collector_Id}${
              obj.Farmers_ID.length < 2
                ? "0"
                : obj.Farmers_ID.length == 2
                ? obj.Farmers_ID
                : obj.Farmers_ID.slice(-2)
            }`
          );
          break;
        case "B":
          cell.value(obj.Respondent_Status);
          break;
        case "C":
          cell.value(
            `${obj.First_Name.trim()} ${obj.Last_Name.trim()} ${
              !obj.Other_Names || obj.Other_Names === "-"
                ? ""
                : obj.Other_Names.trim()
            }`.trim()
          );
          break;
        case "D":
          cell.value(obj.Sex);
          break;
        case "E":
          cell.value("Farm" + obj.FarmStatus.trim().slice(-1));
          break;
        case "F":
          cell.value(obj.Phone_No.trim());
          break;
        case "G":
          for (let i = 0; i < obj.FarmBoundary.split(";").length; i++) {
            const element = obj.FarmBoundary.split(";")[i].trim().split(" ");
            sheet.cell(maxrowNum + i, "G").value(parseFloat(element[0].trim()));
          }
          break;
        case "H":
          for (let i = 0; i < obj.FarmBoundary.split(";").length; i++) {
            const element = obj.FarmBoundary.split(";")[i].trim().split(" ");
            sheet.cell(maxrowNum + i, "H").value(parseFloat(element[1].trim()));
          }
          break;
        case "I":
          cell.value(parseFloat(obj.Total_Area_Ha));
          break;
        case "J":
          cell.value(obj.AreaCropped);
          break;
        case "K":
          planted_Crops[obj._index].forEach((crop, i) => {
            sheet.cell(maxrowNum + i, "K").value(crop);
          });
          break;
        case "L":
          cell.value(obj.Farm_Community);
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
      .row(sheet.usedRange()._maxRowNumber)
      .style({ bottomBorder: { style: "thin", color: "0f0f0f" } });
    // for (const key in sheet._dataValidations) {
    //   if (key[1] === "1") {
    //     const newKey = key + " " + key + maxrowNum;
    //     const value = sheet._dataValidations[key];

    //     value.attributes.sqref = newKey;
    //     sheet._dataValidations[newKey] = value;
    //     delete sheet._dataValidations[key];
    //   }
    // }
    return workbook.toFileAsync("./EnumerationResult.xlsx");
  });
});

// fs.writeFileSync("Enumeration.json", JSON.stringify(result));
