const puppeteer = require('puppeteer');
const XLSX = require('xlsx');
const stringSimilarity = require('string-similarity')
const res = []

async function searchBookPrice() {

  const workbook = XLSX.readFile('books.xlsx');
  const worksheet = workbook.Sheets['Sheet1'];
  const Books = XLSX.utils.sheet_to_json(worksheet);
  console.log(Books);
  console.log('New One: ')
  let i =0;
  const browserSessions = [];
  for (const book of Books) {
  i++;
  const browser = await puppeteer.launch({
    executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    headless: false,
    args: [
      '--disable-cache',
      '--blink-settings=imagesEnabled=false'
    ]}); 

  browserSessions.push(browser);
  const page = await browser.newPage();

  // Intercept and block unnecessary requests
  await page.setRequestInterception(true);
  page.on('request', (request) => {
    if (
      request.resourceType() === 'stylesheet' ||
      request.resourceType() === 'font' ||
      request.resourceType() === 'image'
    ) {
      request.abort();
    } else {
      request.continue();
    }
  });

  // Set user-profile for each session
  const cookies = [
    {
      name: 'user-profile',
      value: `book${i}`,
      domain: 'snapdeal.com',
      path: '/',
      expires: -1,
      httpOnly: false,
      secure: false,
      sameSite: 'Lax',
    },
  ];

  await page.setCookie(...cookies);
  console.log(`The cookie for page${i}: `)
  console.log(cookies);

// Create a user-profile for the page
await page.evaluateOnNewDocument(() => {
  localStorage.setItem('user_profile', JSON.stringify({
    name: 'John Doe',
    email: 'johndoe@example.com',
    age: 30
  }));
});

  // Navigate to snapdeal.com and search for the book by ISBN and title
  await page.goto(`https://www.snapdeal.com/search?keyword=${book.ISBN}+${book.BookTitle}`);


try{
  // Wait for the search results to load
  await page.waitForSelector('.product-tuple-listing');
  // Get the details of each book on the search results page
  const books = await page.$$eval('.product-tuple-listing', (listings,isbn,i,title) => {
    return listings.map(listing => {
      const titleElement = listing.querySelector('.product-title');
      const priceElement = listing.querySelector('.product-price');
      const authorElement = listing.querySelector('.product-author-name');
      const publisherElement = listing.querySelector('.product-publisher');
      const inStockElement = listing.querySelector('.product-out-of-stock');
      const urlElement = listing.querySelector('.product-desc-rating a');

      // Extract the book title from the title element
      const rawTitle = titleElement ? titleElement.textContent.trim() : '';

      // Remove the text enclosed in parentheses, after a colon : or a hyphen -
      const cleanedTitle = rawTitle.replace(/ *\([^)]*\) *| *-[^-]*$/g, '').toLowerCase();

      return {
        No:i,
        BookTitle: cleanedTitle,
        ISBN:isbn,
        Found:'Yes',
        URL: urlElement ? urlElement.href : '',
        Price: priceElement ? parseFloat(priceElement.textContent.replace(/[^0-9\.]+/g, ''))*1000 : Infinity,
        Author: authorElement ? authorElement.textContent.trim() : '',
        Publisher: publisherElement ? publisherElement.textContent.trim() : '',
        InStock: inStockElement ? 'No' : 'Yes'
      };
    });
  },book.ISBN,i,book.BookTitle);

  // Find the book with the minimum price
  const minPrice = Math.min(...books.map(book => book.Price));
  let cheapestBook = books.find(book => book.Price === minPrice);

  if (!cheapestBook) {
    const cheapestBook = { No:i,BookTitle: book.BookTitle, ISBN: book.ISBN, 
      Found: 'No',URL:'',Price:'',Author:'',Publisher:'',InStock:''
  };
    res.push(cheapestBook);
    console.log('Not found');
    await browser.close();
    continue;
  }
    const excelTitle = book.BookTitle.toLowerCase();
    const searchTitle = cheapestBook.BookTitle.toLowerCase();

  // Remove text inside parentheses and after colon or dash
  const excelTitleClean = excelTitle.replace(/\(.*?\)|:.*|-.*|\W/g, '');
  const searchTitleClean = searchTitle.replace(/\(.*?\)|:.*|-.*|\W/g, '');

  // Calculate the string similarity
  const similarity = stringSimilarity.compareTwoStrings(excelTitleClean, searchTitleClean);

  if (similarity >= 0.9) 
  {
    res.push(cheapestBook);
    console.log()
    console.log(cheapestBook);
  } else 
  {

    const cheapestBook = { No:i,BookTitle: book.BookTitle, ISBN: book.ISBN, 
        Found: 'No',URL:'',Price:'',Author:'',Publisher:'',InStock:''
    };
    res.push(cheapestBook);
    console.log()
    console.log(cheapestBook);
  }

  await browser.close();

  }catch{
      // If the book is not found
      const notFoundBook = { No:i,BookTitle: book.BookTitle, ISBN: book.ISBN, 
        Found: 'No',URL:'',Price:'',Author:'',Publisher:'',InStock:''
    };
      res.push(notFoundBook);
      console.log()
      console.log(notFoundBook);
      // Close the browser
      await browser.close();
  }
}
console.log('Final: ')
console.log(browserSessions);

}

function writeToExcel(data) {
  // Create a new workbook
  const workbook = XLSX.utils.book_new();

  // Convert the data to an array of arrays
  const dataArray = data.map(obj => Object.values(obj));

  // Adding a new worksheet to the workbook
  const worksheet = XLSX.utils.aoa_to_sheet(dataArray);

  // Add a title row for each column
  const titleRow = ['No','BookTitle','ISBN','Found','URL','Price','Author','Publisher','InStock'];
  XLSX.utils.sheet_add_aoa(worksheet, [titleRow], { origin: 'A1' });

  // Append the worksheet to the workbook
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

  // Specify the start position as A2
  const start = { r: 1, c: 0 };

  // Add the data to the worksheet from A2 onwards
  XLSX.utils.sheet_add_aoa(worksheet, dataArray, { origin: start });

  // Write the workbook to a file
  XLSX.writeFile(workbook, 'results.xlsx');
}


// Call the searchBookPrice function and pass the writeToExcel function as a callback
searchBookPrice().then(() => {
  writeToExcel(res);
});
