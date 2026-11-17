import { existsSync } from "node:fs";
import { mkdir } from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";
import { sampleProducts } from "./sample-data.js";

const FILE = path.resolve(process.env.DATA_FILE || "data/products.xlsx");
const SHEET = "Products";

// The order of the columns in the sheet. The column NUMBER is the position here, starting at 1.
const COLUMNS = [
  { header: "id", key: "id", width: 6 },
  { header: "name", key: "name", width: 20 },
  { header: "price", key: "price", width: 10 },
  { header: "category", key: "category", width: 16 },
  { header: "available", key: "available", width: 10 },
];

// Create the file from the sample data (or replace it with a fresh one)
export async function resetFile() {
  await mkdir(path.dirname(FILE), { recursive: true });

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(SHEET);
  sheet.columns = COLUMNS;
  sheet.addRows(sampleProducts);
  sheet.getRow(1).font = { bold: true };

  await workbook.xlsx.writeFile(FILE);
}

// Open the file; if it does not exist yet, create it first
async function openSheet() {
  if (!existsSync(FILE)) await resetFile();

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(FILE);
  return { workbook, sheet: workbook.getWorksheet(SHEET) };
}

// One Excel row -> one JavaScript object
function rowToProduct(row) {
  return {
    id: row.getCell(1).value,
    name: row.getCell(2).value,
    price: row.getCell(3).value,
    category: row.getCell(4).value,
    available: row.getCell(5).value,
  };
}

// All the products of a sheet, in file order
function listProducts(sheet) {
  const products = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;                          // skip the header row
    const product = rowToProduct(row);
    if (typeof product.id === "number") products.push(product);   // ignore rows without a numeric id
  });
  return products;
}


// Only ONE operation may touch the file at a time. Every call waits for the previous one.
let queue = Promise.resolve();
function exclusive(task) {
  const result = queue.then(task);
  queue = result.catch(() => {});                               // a failed task must not block the next ones
  return result;
}

// The sheet row that holds a product (or null)
function findRow(sheet, id) {
  let found = null;
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1 && row.getCell(1).value === id) found = row;
  });
  return found;
}

export function readAll() {
  return exclusive(async () => {
    const { sheet } = await openSheet();
    return listProducts(sheet);
  });
}

export function getById(id) {
  return exclusive(async () => {
    const { sheet } = await openSheet();
    return listProducts(sheet).find((product) => product.id === id) || null;
  });
}

export function create({ name, price, category, available }) {
  return exclusive(async () => {
    const { workbook, sheet } = await openSheet();

    const ids = listProducts(sheet).map((product) => product.id);
    const id = ids.length === 0 ? 1 : Math.max(...ids) + 1;     // next free id

    sheet.addRow([id, name, price, category, available]);
    await workbook.xlsx.writeFile(FILE);

    return { id, name, price, category, available };
  });
}

export function update(id, { name, price, category, available }) {
  return exclusive(async () => {
    const { workbook, sheet } = await openSheet();

    const row = findRow(sheet, id);
    if (!row) return null;

    row.getCell(2).value = name;
    row.getCell(3).value = price;
    row.getCell(4).value = category;
    row.getCell(5).value = available;
    await workbook.xlsx.writeFile(FILE);

    return rowToProduct(row);
  });
}

export function remove(id) {
  return exclusive(async () => {
    const { workbook, sheet } = await openSheet();

    const row = findRow(sheet, id);
    if (!row) return false;

    sheet.spliceRows(row.number, 1);
    await workbook.xlsx.writeFile(FILE);
    return true;
  });
}