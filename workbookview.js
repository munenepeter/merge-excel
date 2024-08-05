async function renderExcelTable(filePath, containerId) {
    const response = await fetch(filePath);
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetNames = workbook.SheetNames;
  
    let tabsHtml = '<div class="tabs mb-4">';
    let tablesHtml = '';
  
    // Create tabs
    sheetNames.forEach((sheetName, index) => {
      tabsHtml += `<button class="tab-button ${index === 0 ? 'active' : ''}" data-tab="${index}">${sheetName}</button>`;
    });
    tabsHtml += '</div>';
  
    // Create tables
    sheetNames.forEach((sheetName, index) => {
      const worksheet = workbook.Sheets[sheetName];
      const range = XLSX.utils.decode_range(worksheet['!ref']);
      
      let tableHtml = `<div class="tab-content ${index === 0 ? 'active' : ''}" data-tab-content="${index}">`;
      tableHtml += '<table class="excel-table min-w-full border-collapse border border-gray-300">';
      
      // Add column headers
      tableHtml += '<thead class="bg-gray-100">';
      tableHtml += '<tr>';
      tableHtml += '<th class="border border-gray-300 bg-gray-200 text-gray-800 px-2 py-1"></th>';
      for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex++) {
        tableHtml += `<th class="border border-gray-300 bg-gray-200 text-gray-800 px-2 py-1">${String.fromCharCode(65 + colIndex)}</th>`;
      }
      tableHtml += '</tr>';
      tableHtml += '</thead>';
  
      // Add rows and data
      tableHtml += '<tbody>';
      for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex++) {
        tableHtml += '<tr>';
        for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex++) {
          const cellAddress = { c: colIndex, r: rowIndex };
          const cellRef = XLSX.utils.encode_cell(cellAddress);
          const cell = worksheet[cellRef];
          const cellValue = cell ? cell.v : '';
  
          if (colIndex === range.s.c) {
            tableHtml += `<td class="border border-gray-300 bg-gray-200 text-gray-800 px-2 py-1">${rowIndex + 1}</td>`;
          }
  
          tableHtml += `<td class="border border-gray-300 px-2 py-1 text-nowrap">${cellValue}</td>`;
        }
        tableHtml += '</tr>';
      }
      tableHtml += '</tbody>';
      tableHtml += '</table></div>';
      
      tablesHtml += tableHtml;
    });
  
    const container = document.getElementById(containerId);
    container.innerHTML = tabsHtml + tablesHtml;
  
    // Add tab functionality
    const tabButtons = document.querySelectorAll('.tab-button');
    const tabContents = document.querySelectorAll('.tab-content');
  
    tabButtons.forEach(button => {
      button.addEventListener('click', () => {
        const tabIndex = button.getAttribute('data-tab');
  
        tabButtons.forEach(btn => btn.classList.remove('active'));
        button.classList.add('active');
  
        tabContents.forEach(content => {
          if (content.getAttribute('data-tab-content') === tabIndex) {
            content.classList.add('active');
          } else {
            content.classList.remove('active');
          }
        });
      });
    });
  }
  