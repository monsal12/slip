require("dotenv").config();
const path = require("path");
const express = require("express");
const session = require("express-session");
const { connectDB } = require("./config/db");
const webRoutes = require("./routes/web");

const app = express();
const port = Number(process.env.PORT || 3000);

app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));

app.use(express.urlencoded({ extended: true }));
app.use(
  session({
    secret: process.env.SESSION_SECRET || "slip-session-secret",
    resave: false,
    saveUninitialized: false,
    cookie: {
      maxAge: 1000 * 60 * 60 * 12
    }
  })
);
app.use(express.static(path.resolve("public")));

app.use(webRoutes);

app.use((err, req, res, next) => {
  console.error(err);
  res.status(500).send("Terjadi kesalahan pada server.");
});

async function startServer() {
  await connectDB();
  app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
  });
}

startServer().catch((error) => {
  console.error("Gagal menjalankan server:", error.message);
  process.exit(1);
});
