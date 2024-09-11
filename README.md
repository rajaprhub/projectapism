const XLSX = require('xlsx');
const axios = require('axios');
const fs = require('fs');

// Load the Excel file
const filePath = 'path/to/your/input.xlsx';
const workbook = XLSX.readFile(filePath);
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

// Convert the sheet to JSON
const data = XLSX.utils.sheet_to_json(sheet);

// Define the API endpoint
const apiUrl = 'https://jsonplaceholder.typicode.com/posts'; // Replace with your API

// Function to post data and return responses
const postData = async (data) => {
  try {
    const responses = await Promise.all(
      data.map(async (item) => {
        try {
          const response = await axios.post(apiUrl, item, {
            headers: {
              'Content-Type': 'application/json',
            },
          });
          return { ...item, ...response.data }; // Combine request data with response
        } catch (error) {
          console.error(`Error posting data: ${error.message}`);
          return { ...item, error: error.message }; // Include error message in response
        }
      })
    );
    return responses;
  } catch (error) {
    console.error(`Error processing data: ${error.message}`);
    throw error;
  }
};

// Function to save responses to the same Excel file
const saveToExcel = (responses) => {
  const ws = XLSX.utils.json_to_sheet(responses);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Response Data');
  XLSX.writeFile(wb, 'path/to/your/output.xlsx');
};

// Main function to process the data
const processExcelFile = async () => {
  try {
    const responses = await postData(data);
    saveToExcel(responses);
    console.log('Data processed and saved to Excel successfully.');
  } catch (error) {
    console.error('Error processing Excel file:', error.message);
  }
};

processExcelFile();
