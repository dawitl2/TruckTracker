import { useState } from "react";
import * as XLSX from "xlsx";
import "./ExcelTester.css";

const TARGET_REFERENCE = "FT00211QWBK0";

function ExcelTester() {
  const [resultRow, setResultRow] = useState(null);
  const [status, setStatus] = useState("");

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setStatus("Reading file...");

    const reader = new FileReader();

    reader.onload = (event) => {
      const workbook = XLSX.read(event.target.result, { type: "binary" });

      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      if (!json.length) {
        setStatus("No data found in file");
        return;
      }

      const columns = Object.keys(json[0]);

      // STEP 1: try to find reference-like column
      let referenceColumn = columns.find((col) => {
        const c = col.toLowerCase().replace(/\s/g, "");

        return (
          c.includes("reference") ||
          c.includes("ref") ||
          c.includes("trx") ||
          c.includes("transaction") ||
          c.includes("id")
        );
      });

      let match = null;

      // STEP 2: if reference column found → search only there
      if (referenceColumn) {
        match = json.find((row) => {
          const value = row[referenceColumn];
          return (
            value &&
            value.toString().trim().includes(TARGET_REFERENCE)
          );
        });
      }

      // STEP 3: fallback → scan ALL columns if not found
      if (!match) {
        match = json.find((row) =>
          Object.values(row).some((val) =>
            val &&
            val.toString().includes(TARGET_REFERENCE)
          )
        );
      }

      if (!match) {
        setResultRow(null);
        setStatus("No matching reference found");
        return;
      }

      setResultRow(match);
      setStatus("Match found");
    };

    reader.readAsBinaryString(file);
  };

  return (
    <div className="excel-container">
      <h2>Truck Reference Scanner</h2>

      <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />

      <p>{status}</p>

      {resultRow && (
        <div className="table-wrapper">
          <table>
            <thead>
              <tr>
                {Object.keys(resultRow).map((col) => (
                  <th key={col}>{col}</th>
                ))}
              </tr>
            </thead>

            <tbody>
              <tr>
                {Object.keys(resultRow).map((col) => (
                  <td key={col}>{resultRow[col]}</td>
                ))}
              </tr>
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

export default ExcelTester;