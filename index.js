require("events").EventEmitter.defaultMaxListeners = 100;
const xlsx = require("xlsx");
const stringSimilarity = require("string-similarity");
const puppeteer = require("puppeteer");

// Load the list of books from an excel file
const workbook = xlsx.readFile("input.xlsx");
const worksheet = workbook.SheetNames[0];
const books = xlsx.utils.sheet_to_json(workbook.Sheets[worksheet]);

async function scrap_books(book) {
  // Launch a headless Chrome browser
  const browser = await puppeteer.launch();
  try {
    // Create a new page
    const page = await browser.newPage();
    await page.setDefaultNavigationTimeout(0);
    await page.setDefaultTimeout(0);

    // Navigate to snapdeal.com
    await page.goto("https://www.snapdeal.com/");

    // Type the ISBN code on the search field
    await page.type("#inputValEnter", JSON.stringify(book.ISBN));
    await page.keyboard.press("Enter");

    // Wait for the search results to load
    await page.waitForSelector(".js-product-list");

    // Get the list of search results
    const searchResults = await page.$$eval(
      ".product-tuple-listing .product-tuple-description .product-desc-rating a .product-title",
      (title) => {
        return title.map((t) => t.innerHTML);
      }
    );

    // Check if there are any search results
    if (searchResults.length === 0) {
      console.log(`${book.Book_Title} not found on snapdeal`);
      book.Found = "No";
      book.Site = "NA";
      book.Author = "NA";
      book.Price = "NA";
      book.Publisher = "NA";
      book.In_Stock = "NA";
      return;
    } else {
      // remove -,:,()
      const trimmedSearchResults = searchResults.map((s) => {
        return s
          .split("(")
          .join(",")
          .split("-")
          .join(",")
          .split(":")
          .join(",")
          .split(",")[0]
          .toLowerCase();
      });
      const allResults = [];
      for (let i = 0; i < searchResults.length; i++) {
        allResults.push({
          originalString: searchResults[i],
          trimmedString: trimmedSearchResults[i],
          price: 0,
        });
      }
      // find best matches
      var matches = stringSimilarity.findBestMatch(
        book.Book_Title.toLowerCase(),
        trimmedSearchResults
      );

      // check for matches having more than 90% similarity
      const bestMatches = [];
      for (let i = 0; i < matches["ratings"].length; i++) {
        if (matches["ratings"][i].rating >= 0.9) {
          bestMatches.push(matches["ratings"][i].target);
        }
      }

      if (bestMatches.length === 0) {
        console.log(`${book.Book_Title} not found on snapdeal`);
        book.Found = "No";
        book.Site = "NA";
        book.Author = "NA";
        book.Price = "NA";
        book.Publisher = "NA";
        book.In_Stock = "NA";
        return;
      } else {
        // to get the original titles of all the best matches
        const finalResults = [];
        for (let i = 0; i < allResults.length; i++) {
          for (let j = 0; j < bestMatches.length; j++) {
            if (bestMatches[j] === allResults[i].trimmedString) {
              finalResults.push({
                originalString: allResults[i].originalString,
                trimmedString: allResults[i].trimmedString,
                price: 0,
              });
              break;
            }
          }
        }

        // getting the price of all the final results to compare
        // Click on the book to open it
        const allPrices = [];
        const resultsWithPrice = [];
        for (const oneBook of finalResults) {
          const pageTarget = page.target(); //save this to know that this was the opener
          await page.evaluate(
            ({ oneBook }) => {
              document
                .querySelector(
                  '.product-tuple-listing .product-tuple-image a .picture-elem .product-image[title="' +
                    oneBook.originalString +
                    '"]'
                )
                .click();
            },
            { oneBook }
          );
          const newTarget = await browser.waitForTarget(
            (target) => target.opener() === pageTarget
          ); //check that you opened this page, rather than just checking the url
          const newPage = await newTarget.page(); //get the page object
          await newPage.waitForSelector(".disp-table"); //wait for page to be loaded
          const info = await newPage.evaluate(() => {
            return document.querySelector(
              ".disp-table .disp-table-cell .pdp-final-price .payBlkBig"
            ).textContent;
          });
          resultsWithPrice.push({
            originalString: oneBook.originalString,
            trimmedString: oneBook.trimmedString,
            price: info,
          });
          allPrices.push(info);
          await newPage.close();
        }
        const minPrice = Math.min(...allPrices);
        const cheapestBook = [];
        for (let k = 0; k < resultsWithPrice.length; k++) {
          if (resultsWithPrice[k].price == parseInt(minPrice)) {
            cheapestBook.push(resultsWithPrice[k]);
          }
        }
        for (let f = 0; f < 1; f++) {
          const pageTarget = page.target(); //save this to know that this was the opener
          await page.evaluate(
            ({ cheapestBook }) => {
              document
                .querySelector(
                  '.product-tuple-listing .product-tuple-image a .picture-elem .product-image[title="' +
                    cheapestBook[0].originalString +
                    '"]'
                )
                .click();
            },
            { cheapestBook }
          );
          const newTarget = await browser.waitForTarget(
            (target) => target.opener() === pageTarget
          ); //check that you opened this page, rather than just checking the url
          const newPage = await newTarget.page(); //get the page object
          await newPage.waitForSelector(".disp-table"); //wait for page to be loaded
          book.Price = cheapestBook[0].price;
          book.Found = "Yes";
          book.Author = await newPage.$eval(
            ".p-keyfeatures .clearfix :nth-child(5) .h-content",
            (elem) => elem.innerText.split(":")[1]
          );
          book.Publisher = await newPage.$eval(
            ".p-keyfeatures .clearfix :nth-child(3) .h-content",
            (elem) => elem.innerText.split(":")[1]
          );
          book.In_Stock = "Yes";
          book.Site = newPage.url();
        }
      }
    }
  } catch (err) {
    console.error(err);
  } finally {
    // Close the browser
    await browser.close();
    var bookWS = xlsx.utils.json_to_sheet(books);
    var wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, bookWS, "sheet1");
    xlsx.writeFile(wb, "Output.xlsx");
  }
}

// Loop through each book and search for it on snapdeal.com
for (let i = 0; i < books.length; i++) {
  scrap_books(books[i]);
}
