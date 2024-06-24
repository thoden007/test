

import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import './App.css';

function App() {
  const [file, setFile] = useState(null);
  const [uniqueValues, setUniqueValues] = useState([]);
  const [error, setError] = useState('');

  const handleFileUpload = (e) => {
    setFile(e.target.files[0]);
    setUniqueValues([]);
    setError('');
  };

  const handleRun = () => {
    if (file) {
      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const data = new Uint8Array(event.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const sheetName = workbook.SheetNames[0]; // Select the second sheet (index 1)
          if (sheetName) {
            const worksheet = workbook.Sheets[sheetName];
            const json = XLSX.utils.sheet_to_json(worksheet);
            findUniqueValues(json);
          } else {
            setError("Sheet 2 not found in the workbook.");
          }
        } catch (error) {
          setError("Error reading the file. Please ensure it's a valid Excel file.");
        }
      };
      reader.readAsArrayBuffer(file);
    }
  };

  const findUniqueValues = (json) => {
    const values = json.flatMap(Object.values);
    const lowerCaseValues = values.map(value => value && value.toString().toLowerCase());

    const uniqueValuesSet = new Set(lowerCaseValues);
    const uniqueValuesArray = Array.from(uniqueValuesSet);

    const uniqueOriginalValues = uniqueValuesArray.map(lowerCaseValue => {
      const index = lowerCaseValues.indexOf(lowerCaseValue);
      return values[index];
    });

    const sortedUniqueValues = uniqueOriginalValues.sort((a, b) => {
      if (a && b) {
        return a.toString().localeCompare(b.toString());
      }
      return 0;
    });

    setUniqueValues(sortedUniqueValues);
  };

  const handleExport = () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(uniqueValues.map(value => ({ Unique: value })));
    XLSX.utils.book_append_sheet(wb, ws, 'UniqueValues');
    XLSX.writeFile(wb, 'unique_values.xlsx');
  };

  const handleDownloadUniqueValues = () => {
    if (uniqueValues.length > 0) {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(uniqueValues.map(value => ({ Unique: value })));
      XLSX.utils.book_append_sheet(wb, ws, 'UniqueValues');
      XLSX.writeFile(wb, 'unique_values.xlsx');
    } else {
      setError("No unique values found to download.");
    }
  };

  return (
    <div className="A">
      <h1>Upload Excel File </h1>
      <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
      <button onClick={handleRun} disabled={!file}>Run</button>
      {error && <div className="error">{error}</div>}
      {uniqueValues.length > 0 && (
        <div>
          <h2>Unique Values:</h2>
          <p>no1</p>
          <p>Total Unique Values Found: {uniqueValues.length}</p>
          <ul>
            {uniqueValues.map((value, index) => (
              <li key={index}>{value}</li>
            ))}
          </ul>
          <button onClick={handleExport}>Export to Excel</button>
          <button onClick={handleDownloadUniqueValues}>Download Unique Values</button>
        </div>
      )}
    </div>
  );
}

export default App;
