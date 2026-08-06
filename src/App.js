import React, { useState } from "react";
import * as XLSX from "xlsx";
import "./App.css";

function App() {

  const [data, setData] = useState([]);
  const [file, setFile] = useState(null);
  const [processed, setProcessed] = useState([]);
  const [index, setIndex] = useState(0);

  const [verifiedCount, setVerifiedCount] = useState("");
  const [remarks, setRemarks] = useState("");

  // ✅ Upload Excel
  const handleFileUpload = (e) => {
  const uploadedFile = e.target.files[0];
  setFile(uploadedFile);

  const reader = new FileReader();

  reader.onload = (evt) => {
    const wb = XLSX.read(evt.target.result, { type: "binary" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);

    setData(json);
    setProcessed(json);
    setIndex(0);
  };

  reader.readAsBinaryString(uploadedFile);
};

  const current = data[index];

  // ✅ Entity Name
  const getEntityName = () => {
    try {
      const payload = JSON.parse(current["Payload"]);
      return payload?.name || "N/A";
    } catch {
      return "N/A";
    }
  };

  // ✅ Dynamic source detection (best fix)
  const getSourceColumns = () => {

    if (!current) return [];

    const keys = Object.keys(current);

    const excludeCols = [
      "Entity ID",
      "Client Name",
      "Industry ID",
      "Sync Partner",
      "Client Products",
      "Entity Type",
      "Payload",
      "Status",
      "Verified URL Count",
      "Remarks"
    ];

    return keys.filter(col =>
      !excludeCols.includes(col)
    );
  };

  // ✅ Error table logic
  const getErrors = () => {

  if (!current) return [];

  const excludeCols = [
    "client id",
    "entity created date",
    "modified time",
    "code",
    "last run",
    "source of truth",
    "verification needed",
    "payload",
    "entity id",
    "client name",
    "industry id",
    "sync partner",
    "client products",
    "entity type",
    "status",
    "verified url count",
    "remarks"
  ];

  return Object.keys(current)
    .filter(col => {

      if (!col) return false;

      const lower = col.toLowerCase();

      // ✅ exclude unwanted fields
      return !excludeCols.includes(lower);

    })
    .map(col => {

      const val = current[col];

      if (!val || val.toString().trim() === "") return null;

      const strVal = val.toString();

      let action = "";

      if (strVal.includes("404") || strVal.includes("No URLs")) {
        action = "Add URL";
      } else if (strVal.includes("500")) {
        action = "Retry";
      } else if (strVal.includes("Different URL already exists")) {
        action = "Verify Duplicate";
      }

      return {
        source: col,
        value: strVal,
        action
      };
    })
    .filter(x => x !== null);
};
  // ✅ Update row
  const handleUpdate = (status) => {

    let updated = [...processed];

    updated[index] = {
      ...updated[index],
      Status: status,
      "Verified URL Count": Number(verifiedCount),
      Remarks: remarks
    };

    setProcessed(updated);
    setIndex(index + 1);

    setVerifiedCount("");
    setRemarks("");
  };

  // ✅ Summary
  const total = processed.length;
  const completed = processed.filter(d => d.Status === "Completed").length;
  const pending = processed.filter(d => d.Status === "Pending").length;

  const totalURLs = processed.reduce(
    (sum, d) => sum + Number(d["Verified URL Count"] || 0),
    0
  );

  // ✅ Download Excel (fixed placement logic)
 const downloadExcel = () => {

  if (!file) return;

  const reader = new FileReader();

  reader.onload = (e) => {

    const wb = XLSX.read(e.target.result, { type: "binary" });
    const ws = wb.Sheets[wb.SheetNames[0]];

    // ✅ Read as array (preserve structure)
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

    const headers = data[0];

    // ✅ Add new headers
    headers.push("Status", "Verified URL Count", "Remarks");

    const totalCols = headers.length;

    // ✅ Normalize ALL rows to same length
    for (let i = 1; i < data.length; i++) {

      // ✅ Fill missing cells BEFORE pushing
      while (data[i].length < totalCols - 3) {
        data[i].push(null);
      }

      const rowData = processed[i - 1] || {};

      data[i].push(
        rowData.Status || "",
        rowData["Verified URL Count"] || "",
        rowData.Remarks || ""
      );
    }

    const newWs = XLSX.utils.aoa_to_sheet(data);
    const newWb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(newWb, newWs, "Output");

    XLSX.writeFile(newWb, "processed_output.xlsx");
  };

  reader.readAsBinaryString(file);
};
  return (
    <div className="container">

  <div className="header">🚀 URL Discovery Processing Tool</div>

  <input type="file" onChange={handleFileUpload} />

  {/* ✅ Summary */}
  <div className="summary">
    <div className="card">Total Entities<br/><b>{total}</b></div>
    <div className="card">Completed<br/><b>{completed}</b></div>
    <div className="card">Pending<br/><b>{pending}</b></div>
    <div className="card">Total URLs Verified<br/><b>{totalURLs}</b></div>
  </div>

  {/* ✅ Progress */}
  {total > 0 && (
    <>
      <progress value={index} max={total}></progress>
      <p>{index} / {total}</p>
    </>
  )}

  {/* ✅ ENTITY PANEL */}
  {current && (
    <div className="main-card">

      <div className="entity-section">

  <div className="entity-title">
    {getEntityName()}
  </div>

  <div className="entity-id">
    <a
      href={`https://crawler-admin.consumerism.pressganey.com/#/verify-sources/${current["Entity ID"]}`}
      target="_blank"
      rel="noreferrer"
    >
      {current["Entity ID"]}
    </a>
  </div>

  <div className="entity-grid">

    <div className="info-card">
      <span>Client</span>
      <b>{current["Client Name"]}</b>
    </div>

    <div className="info-card">
      <span>Sync Type</span>
      <b>{current["Sync Partner"]}</b>
    </div>

    <div className="info-card">
      <span>Products</span>
      <b>{current["Client Products"]}</b>
    </div>

    <div className="info-card">
      <span>Industry</span>
      <b>{current["Industry ID"]}</b>
    </div>

    <div className="info-card">
      <span>Entity Type</span>
      <b>{current["Entity Type"]}</b>
    </div>

  </div>

</div>


      {/* ✅ Error Table */}
      <div className="section-title">🔍 Error Details</div>

      <table>
        <thead>
          <tr>
            <th>Source</th>
            <th>Value</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody>
          {getErrors().map((e, i) => (
            <tr key={i}>
              <td>{e.source}</td>
              <td>{e.value}</td>
              <td className={e.action ? "error" : ""}>{e.action}</td>
            </tr>
          ))}
        </tbody>
      </table>

      {/* ✅ Input Section */}
      <div className="section-title">✅ Verification-new</div>

      <input
        placeholder="No of URLs Verified"
        value={verifiedCount}
        onChange={(e) => setVerifiedCount(e.target.value)}
      />

      <textarea
        placeholder="Remarks"
        value={remarks}
        onChange={(e) => setRemarks(e.target.value)}
      />

      {/* ✅ Action Buttons */}
      {verifiedCount && (
        <div>
          <button className="btn btn-success" onClick={() => handleUpdate("Completed")}>
            ✅ Completed
          </button>

          <button className="btn btn-warning" onClick={() => handleUpdate("Pending")}>
            ⏳ Pending
          </button>
        </div>
      )}

    </div>
  )}

  {/* ✅ Download */}
  {processed.length > 0 && (
    <button className="btn btn-download" onClick={downloadExcel}>
      ⬇ Download Output File
    </button>
  )}

</div>

  );
}

export default App;
