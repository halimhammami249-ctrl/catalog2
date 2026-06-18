const XLSX = require('xlsx');
const fs = require('fs-extra');
const path = require('path');

// =====================
// CONFIG
// =====================
const EXCEL_FILE = '../importer/products.xlsx';
const INPUT_IMAGES = '../importer/images';

const OUTPUT_DIR = '../public';
const OUTPUT_JSON = path.join(OUTPUT_DIR, '../public/products.json');
const OUTPUT_IMAGES = path.join(OUTPUT_DIR, '../public/images');

// =====================
// CREATE OUTPUT FOLDERS
// =====================
fs.ensureDirSync(OUTPUT_DIR);
fs.ensureDirSync(OUTPUT_IMAGES);

// =====================
// READ EXCEL
// =====================
console.log('📥 Starting import...');

const workbook = XLSX.readFile(EXCEL_FILE);
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const rows = XLSX.utils.sheet_to_json(sheet);

// =====================
// PROCESS PRODUCTS
// =====================
const products = [];

let success = 0;
let missingImages = 0;

rows.forEach((row) => {
  console.log('Processing:', row.name);

  const imageName = row.image;
  const sourceImage = path.join(INPUT_IMAGES, imageName);
  const targetImage = path.join(OUTPUT_IMAGES, imageName);

  // Copy image if exists
  if (fs.existsSync(sourceImage)) {
    fs.copySync(sourceImage, targetImage);
  } else {
    console.log('⚠️ Missing image:', imageName);
    missingImages++;
  }

  products.push({
    id: row.id,
    name: row.name,
    price: row.price,
    image: `/images/${imageName}`,
    description: row.description || '',
  });

  success++;
});

// =====================
// WRITE JSON FILE
// =====================
fs.writeFileSync(OUTPUT_JSON, JSON.stringify(products, null, 2));

// =====================
// SUMMARY
// =====================
console.log('\n✅ Import completed');
console.log('✔ Products imported:', success);
console.log('⚠️ Missing images:', missingImages);
console.log('📦 Output: public/products.json + public/images/');
