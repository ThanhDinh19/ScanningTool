import app from "./app.js";
import { resetGrandTotalExcel } from "./utils/math.util.js";

const PORT = 5000;

resetGrandTotalExcel();

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
