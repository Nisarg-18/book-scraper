const xlsx = require("xlsx");

exports.writeInExcel = (books) => {
  var bookWS = xlsx.utils.json_to_sheet(books);
  var wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, bookWS, "sheet1");
  xlsx.writeFile(wb, "Output.xlsx");
};
