import * as XLSX from 'xlsx';

export const parseExcelHeaders = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Convert sheet to array of arrays to easily scan rows
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });

        if (rows.length === 0) {
          throw new Error("The Excel sheet is empty.");
        }

        let headerRowIndex = 0;
        let maxValidCells = 0;

        // Scan the first 50 rows to find the one with the most non-empty strings.
        // Bank reports often have titles and metadata in the first few rows.
        for (let i = 0; i < Math.min(rows.length, 50); i++) {
          const row = rows[i] || [];
          const validCells = row.filter(cell => cell !== null && cell !== undefined && String(cell).trim() !== '').length;

          if (validCells > maxValidCells) {
            maxValidCells = validCells;
            headerRowIndex = i;
          }
        }

        const rawHeaders = rows[headerRowIndex] || [];
        const headers = [];
        const validIndices = [];

        // Collect valid headers
        rawHeaders.forEach((header, index) => {
          if (header !== null && header !== undefined && String(header).trim() !== '') {
            // Check for duplicates and append index if needed
            let headerName = String(header).trim();
            if (headers.includes(headerName)) {
              headerName = `${headerName}_${index}`;
            }
            headers.push(headerName);
            validIndices.push(index);
          }
        });

        // Parse data starting from the row AFTER the header row
        const jsonData = [];
        for (let i = headerRowIndex + 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row) continue;

          const isRowEmpty = row.every(cell => cell === null || cell === undefined || String(cell).trim() === '');
          if (isRowEmpty) continue;

          const rowObj = {};
          validIndices.forEach((colIndex, idx) => {
            const headerName = headers[idx];
            rowObj[headerName] = row[colIndex];
          });
          jsonData.push(rowObj);
        }

        resolve({ headers, rawData: jsonData });
      } catch (error) {
        console.error("Error parsing excel:", error);
        reject(error);
      }
    };
    reader.onerror = (error) => reject(error);
    reader.readAsArrayBuffer(file);
  });
};

export const parseMappingFile = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Simple 2-column mapping parsing
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });

        const mappingDictionary = {};

        let headerPassed = false;

        rows.forEach(row => {
          if (!row || row.length < 2) return;

          const code = String(row[0] || '').trim();
          const loanType = String(row[1] || '').trim();

          if (!code || !loanType) return;

          // Skip header row if it resembles 'code' and 'type' heavily or is the first valid row and we haven't seen one yet
          if (!headerPassed) {
            headerPassed = true; // Assume first valid row might be header
            if (code.toLowerCase().includes('code') || code.toLowerCase().includes('facility')) {
              return;
            }
          }

          mappingDictionary[code] = loanType;
        });

        resolve(mappingDictionary);
      } catch (error) {
        console.error("Error parsing mapping file:", error);
        reject(error);
      }
    };
    reader.onerror = (error) => reject(error);
    reader.readAsArrayBuffer(file);
  });
};
