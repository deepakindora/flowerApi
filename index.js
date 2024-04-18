const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const basicAuth = require('express-basic-auth');
const { createProxyMiddleware } = require('http-proxy-middleware');

const app = express();
const PORT = 8080;

app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', 'http://localhost:8080');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  res.setHeader('Access-Control-Allow-Credentials', 'true');

  next();
});

app.use(bodyParser.json());

// Define basic authentication credentials
const users = { 'admin': 'password123' };

// Enable basic authentication for all APIs
app.use(basicAuth({ users, challenge: true, unauthorizedResponse: 'Unauthorized' }));

const excelFilePath = './db/messages.xlsx';

// Load existing messages from the Excel file
let messages = loadMessagesFromExcel();

// API to create a post with name, email, and message
app.post('/queryMessage', (req, res) => {
  const { name, email, message } = req.body;
  const messageID = messages.length + 1;

  const newMessage = {
    messageID,
    name,
    email,
    message,
  };

  messages.push(newMessage);

  // Update the Excel file with the new data
  updateExcelFile(messages);

  res.status(201).json({ message: 'Message created successfully', messageID });
});

// API to get all data stored by queryMessage
app.get('/messageData', (req, res) => {
  res.json(messages);
});

// API to get details based on messageID
app.get('/messageDetail/:messageID', (req, res) => {
  const { messageID } = req.params;
  const message = messages.find((msg) => msg.messageID == messageID);

  if (message) {
    res.json(message);
  } else {
    res.status(404).json({ error: 'Message not found' });
  }
});

// API to delete a message based on messageID
app.delete('/deleteMessage/:messageID', (req, res) => {
    const { messageID } = req.params;
    const index = messages.findIndex((msg) => msg.messageID == messageID);
  
    if (index !== -1) {
      messages.splice(index, 1);
      updateExcelFile(messages);
      res.json({ message: 'Message deleted successfully' });
    } else {
      res.status(404).json({ error: 'Message not found' });
    }
  });

  // Configure proxy middleware for other services
// const apiProxy = createProxyMiddleware('/jqueryAjax', {
//   target: 'http://localhost:8080/jqueryAjax.html',
//   changeOrigin: true,
// });

// app.use('jqueryAjax', apiProxy);

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});

function loadMessagesFromExcel() {
    try {
      const fs = require('fs');
  
      // Check if the file exists before attempting to read it
      if (fs.existsSync(excelFilePath)) {
        const workbook = xlsx.readFile(excelFilePath);
        const sheetName = workbook.SheetNames[0];
        const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
        return sheetData || [];
      } else {
        console.warn('Excel file does not exist. Returning empty array.');
        return [];
      }
    } catch (error) {
      console.error('Error loading messages from Excel:', error);
      return [];
    }
  }
  

function updateExcelFile(data) {
  const ws = xlsx.utils.json_to_sheet(data);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, 'Messages');
  xlsx.writeFile(wb, excelFilePath);
}
