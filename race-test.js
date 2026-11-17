import { create, readAll, resetFile } from "./excelStore.js";

await resetFile();                                              // start from the 4 sample products

// Ten creates started AT THE SAME TIME, like ten users clicking "Add" together
await Promise.all(
  Array.from({ length: 10 }, (_, index) =>
    create({ name: `Item ${index + 1}`, price: index + 1, category: "Test", available: true })
  )
);

const products = await readAll();
console.log("Expected 14 products, found", products.length);
console.log("ids:", products.map((product) => product.id).join(", "));