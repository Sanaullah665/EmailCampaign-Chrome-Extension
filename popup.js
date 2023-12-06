document.getElementById('file-input').addEventListener('change', handleFileSelect);

let emailList = []; // Store emails globally

function handleFileSelect(event) {
  const file = event.target.files[0];
  if (file) {
    parseExcel(file);
  }
}

function parseExcel(file) {
  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      // Assuming the first sheet is relevant, you can modify this if needed
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      // Assuming the header is in the first row
      const header = XLSX.utils.sheet_to_json(sheet, { header: 1 })[0];

      if (!header.includes("email")) {
        showError("No 'email' header found in the file.");
        return;
      }

      emailList = XLSX.utils.sheet_to_json(sheet, { header: "email" });
      console.log("Parsed emails:", emailList);
      displayEmailList(emailList);
    } catch (error) {
      console.error("Error parsing the Excel file:", error);
      showError("Error parsing the Excel file. Please make sure it's a valid Excel file.");
    }
  };

  reader.readAsArrayBuffer(file);
}

function displayEmailList(emails) {
  console.log("Displaying emails:", emails);
  const emailListDiv = document.getElementById('email-list');
  emailListDiv.innerHTML = '<h3>Email List:</h3>';

  if (emails.length === 0) {
    showError("No emails found in the file.");
    return;
  }

  const ul = document.createElement('ul');
  emails.forEach(email => {
    const li = document.createElement('li');
    li.textContent = email;
    ul.appendChild(li);
  });

  emailListDiv.appendChild(ul);
}

function sendMessage() {
  const message = document.getElementById('message-box').value;

  if (!message.trim()) {
    showError("Please compose a message before sending.");
    return;
  }

  if (emailList.length === 0) {
    showError("No emails available to send the message to.");
    return;
  }

  console.log("Sending message:", message, "to emails:", emailList);

  // TODO: Implement your email sending logic here
  // For simplicity, we'll just display a success message
  const statusDiv = document.getElementById('status');
  statusDiv.textContent = "Message sent successfully!";
  statusDiv.style.color = '#2ecc71';
}

function showError(message) {
  console.error("Error:", message);
  const statusDiv = document.getElementById('status');
  statusDiv.textContent = message;
  statusDiv.style.color = '#e74c3c';
}
// Attach functions to the window object for accessibility in the HTML
window.handleFileSelect = handleFileSelect;
window.sendMessage = sendMessage;