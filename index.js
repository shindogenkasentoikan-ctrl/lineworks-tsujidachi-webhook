import express from "express";

const app = express();
const PORT = process.env.PORT || 8080;

app.get("/", (_req, res) => {
  res.status(200).send("cloud run is alive minimal");
});

app.post("/", (_req, res) => {
  res.status(200).json({ ok: true });
});

app.listen(PORT, "0.0.0.0", () => {
  console.log(`Listening on port ${PORT}`);
});
