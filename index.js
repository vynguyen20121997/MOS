const ExcelJS = require("exceljs");
const _ = require("lodash");

async function readExcelFile() {
  try {
    // Tạo workbook mới
    const baiLam = new ExcelJS.Workbook();
    const dapAn = new ExcelJS.Workbook();

    // Đọc file Excel từ đường dẫn cụ thể
    await baiLam.xlsx.readFile(
      "C:\\Users\\Admin\\Desktop\\Project1\\test.xlsx"
    );

    await dapAn.xlsx.readFile(
      "C:\\Users\\Admin\\Desktop\\Project1\\File_Dapan.xlsx"
    );

    // So sánh tất cả các properties trong baiLam và dapAn
    const compareSheetsValues = (sheet1, sheet2) => {
      const sheet1Data = sheet1.getSheetValues();
      const sheet2Data = sheet2.getSheetValues();
      return _.isEqual(sheet1Data, sheet2Data);
    };

    const compareProperties = (obj1, obj2) => {
      return _.isEqual(obj1, obj2);
    };

    const results = baiLam.worksheets.map((sheet, index) => {
      const dapAnSheet = dapAn.worksheets[index];
      if (dapAnSheet) {
        const sheetComparison = compareSheetsValues(sheet, dapAnSheet);
        const propertiesComparison = compareProperties(sheet, dapAnSheet);
        if (!propertiesComparison) {
          const differences = _.reduce(
            sheet,
            (result, value, key) => {
              if (key !== "_workbook" && !_.isEqual(value, dapAnSheet[key])) {
                const propertyDifferences = _.reduce(
                  value,
                  (propResult, propValue, propKey) => {
                    const dapAnSheetValue = dapAnSheet[key][propKey];
                    if (!_.isEqual(propValue, dapAnSheetValue)) {
                      propResult.push({
                        property: propKey,
                        value1: propValue,
                        value2: dapAnSheet[key][propKey],
                      });
                    }
                    return propResult;
                  },
                  []
                );
                result.push({ key, differences: propertyDifferences });
              }
              return result;
            },
            []
          );
          console.log(`Differences in sheet ${index + 1}:`, differences);
        }
        return sheetComparison && propertiesComparison;
      }
      return false;
    });

    console.log("Kết quả so sánh các sheet:", results);

    // Lấy danh sách tất cả các worksheets
    // console.log("Danh sách các sheet trong file:");
    // baiLam.worksheets.forEach((sheet, index) => {
    //   console.log(`${index + 1}. ${sheet.name}`);
    // });

    // Lấy worksheet đầu tiên
    // const worksheet = workbook1.getWorksheet(1);

    // Mảng để lưu dữ liệu
    // const data = [];

    // Đọc từng hàng
    // worksheet.eachRow((row, rowNumber) => {
    //   const rowData = [];
    //   row.eachCell((cell) => {
    //     rowData.push(cell.value);
    //   });
    //   data.push(rowData);

    //   // In ra dữ liệu từng hàng để kiểm tra
    //   console.log(`Hàng ${rowNumber}:`, rowData);
    // });

    return data;
  } catch (error) {
    if (error.code === "ENOENT") {
      console.error(
        "Không tìm thấy file Excel. Vui lòng kiểm tra lại đường dẫn."
      );
    } else {
      console.error("Lỗi khi đọc file Excel:", error.message);
    }
    throw error;
  }
}

// Chạy hàm đọc file
readExcelFile()
  .then((data) => {
    console.log("\nTổng số hàng đã đọc:", data.length);
  })
  .catch((error) => {
    console.log("Đã xảy ra lỗi khi đọc file");
  });
