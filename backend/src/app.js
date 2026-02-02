import express from "express";
import cors from "cors";
import dataRoute from "./routes/excel.route.js";

const app = express();

app.use(cors());
app.use(express.json());

app.use("/api", dataRoute); 

export default app;
