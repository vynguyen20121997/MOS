const ExcelJS = require("exceljs");

async function readExcelFile() {
  try {
    // Tạo workbook mới
    const workbook = new ExcelJS.Workbook();

    // Đọc file Excel từ đường dẫn cụ thể
    await workbook.xlsx.readFile(
      "C:\\Users\\Admin\\Desktop\\Project1\\test.xlsx"
    );

    // Lấy danh sách tất cả các worksheets
    console.log("Danh sách các sheet trong file:");
    workbook.worksheets.forEach((sheet, index) => {
      console.log(`${index + 1}. ${sheet.name}`);
    });

    // Lấy worksheet đầu tiên
    const worksheet = workbook.getWorksheet(1);
    console.log(`\nĐang đọc sheet: ${worksheet.name}`);

    // Mảng để lưu dữ liệu
    const data = [];

    // Đọc từng hàng
    worksheet.eachRow((row, rowNumber) => {
      const rowData = [];
      row.eachCell((cell) => {
        rowData.push(cell.value);
      });
      data.push(rowData);

      // In ra dữ liệu từng hàng để kiểm tra
      console.log(`Hàng ${rowNumber}:`, rowData);
    });

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
