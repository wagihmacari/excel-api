// server.js: a REST API on top of the Excel store.
import express from "express";
import cors from "cors";
import * as store from "./excelStore.js";

const app = express();
const PORT = process.env.PORT || 3000;

// Which web pages may call this API? Set ALLOWED_ORIGIN on the server, for example
//   ALLOWED_ORIGIN=https://myname.github.io
// Several origins: separate them with commas. Not set = anybody (fine for local tests).
const allowedOrigins = (process.env.ALLOWED_ORIGIN || "*").split(",").map((origin) => origin.trim());

app.use(cors({ origin: allowedOrigins.includes("*") ? "*" : allowedOrigins }));
app.use(express.json());        // read JSON bodies

// Log every request
app.use((request, response, next) => {
  console.log(`${request.method} ${request.url}`);
  next();
});

app.get("/", (request, response) => {
  response.json({ message: "Excel API is running", products: "/api/products", download: "/api/export" });
});

// Used by hosting platforms (and by you) to check that the server is alive
app.get("/health", (request, response) => {
  response.json({ status: "ok" });
});

// ===== Read =====

// GET /api/products          -> all products
app.get("/api/products", async (request, response) => {
  response.json(await store.readAll());
});

// GET /api/products/2        -> one product
app.get("/api/products/:id", async (request, response) => {
  const id = Number(request.params.id);               // route parameters are always text
  const product = await store.getById(id);

  if (!product) {
    return response.status(404).json({ message: `Product ${id} not found` });
  }
  response.json(product);
});

// ===== Validation: never trust what the client sends =====

// Returns { error } when the data is wrong, or { product } with clean values
function validateProduct(body) {
  const { name, price, category, available } = body ?? {};       // body is undefined if no JSON was sent

  if (typeof name !== "string" || name.trim() === "") {
    return { error: "name must be a non-empty text" };
  }
  if (typeof price !== "number" || !Number.isFinite(price) || price < 0) {
    return { error: "price must be a number, 0 or more" };
  }
  if (category !== undefined && typeof category !== "string") {
    return { error: "category must be text" };
  }
  if (available !== undefined && typeof available !== "boolean") {
    return { error: "available must be true or false" };
  }

  return {
    product: {
      name: name.trim(),
      price,
      category: (category || "Other").trim(),
      available: available ?? true,
    },
  };
}

// ===== Create =====

// POST /api/products         -> add a product (JSON body)
app.post("/api/products", async (request, response) => {
  const { error, product } = validateProduct(request.body);
  if (error) {
    return response.status(400).json({ message: error });
  }

  const created = await store.create(product);
  response.status(201).json(created);                   // 201 Created
});

// ===== Update =====

// PUT /api/products/2        -> replace the data of product 2 (JSON body)
app.put("/api/products/:id", async (request, response) => {
  const id = Number(request.params.id);

  const { error, product } = validateProduct(request.body);
  if (error) {
    return response.status(400).json({ message: error });
  }

  const updated = await store.update(id, product);
  if (!updated) {
    return response.status(404).json({ message: `Product ${id} not found` });
  }
  response.json(updated);
});

// ===== Delete =====

// DELETE /api/products/2     -> remove product 2
app.delete("/api/products/:id", async (request, response) => {
  const id = Number(request.params.id);

  const removed = await store.remove(id);
  if (!removed) {
    return response.status(404).json({ message: `Product ${id} not found` });
  }
  response.sendStatus(204);                             // 204 No Content: done, nothing to send back
});

// ===== Download the Excel file =====

// GET /api/export            -> the current products.xlsx
app.get("/api/export", async (request, response) => {
  const buffer = await store.exportBuffer();
  response.set({
    "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "Content-Disposition": 'attachment; filename="products.xlsx"',
  });
  response.send(buffer);
});

// ===== After all routes =====

// No route matched
app.use((request, response) => {
  response.status(404).json({ message: "Not found" });
});

// Error handler: Express recognises it by its FOUR parameters
app.use((error, request, response, next) => {
  if (error.type === "entity.parse.failed") {
    return response.status(400).json({ message: "The body is not valid JSON" });
  }
  console.error(error);
  response.status(500).json({ message: "Something went wrong on the server" });
});

// 0.0.0.0 = accept connections from outside the computer (needed on Render)
app.listen(PORT, "0.0.0.0", () => {
  console.log(`Excel API running at http://localhost:${PORT}`);
});
export function exportBuffer() {
  return exclusive(async () => {
    const { workbook } = await openSheet();
    return workbook.xlsx.writeBuffer();
  });
}