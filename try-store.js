import { readAll, getById, create, update, remove, resetFile } from "./excelStore.js";

await resetFile();                                              // start from the 4 sample products

console.log("readAll():");
console.table(await readAll());

console.log("getById(2):", await getById(2));
console.log("getById(99):", await getById(99));

const lamp = await create({ name: "Lamp", price: 45, category: "Furniture", available: true });
console.log("create():", lamp);

console.log("update(5):", await update(5, { name: "Big lamp", price: 60, category: "Furniture", available: false }));
console.log("update(99):", await update(99, { name: "Ghost", price: 1, category: "Other", available: true }));

console.log("remove(1):", await remove(1));
console.log("remove(99):", await remove(99));

console.log("readAll() at the end:");
console.table(await readAll());