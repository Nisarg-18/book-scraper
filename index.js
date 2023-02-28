require("events").EventEmitter.defaultMaxListeners = 100;
const xlsx = require("xlsx");
const { scrap_books } = require("./helper/scrapeBooks");

// Load the list of books from an excel file
const workbook = xlsx.readFile("input.xlsx");
const worksheet = workbook.SheetNames[0];
const books = xlsx.utils.sheet_to_json(workbook.Sheets[worksheet]);

// Loop through each book and search for it on snapdeal.com
for (let i = 0; i < books.length; i++) {
  scrap_books(books[i], books);
}
