import app from "./app.js";
import { resetGrandTotalExcel } from "./utils/math.util.js";

const PORT = 5000;

resetGrandTotalExcel();

app.listen(5000, "0.0.0.0", () => {
  console.log("Server running on 0.0.0.0:5000");
});