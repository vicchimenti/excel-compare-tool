// Store the selected files
let selectedFiles = [];

// Add event listener for the file selection button
document.getElementById('select-files-btn').addEventListener('click', async () => {
  try {
    selectedFiles = await window.api.selectFiles();
    
    if (selectedFiles.length >= 2) {
      document.getElementById('file1-path').textContent = selectedFiles[0];
      document.getElementById('file2-path').textContent = selectedFiles[1];
      document.getElementById('compare-btn').disabled = false;
    } else {
      alert('Please select at least 2 Excel files to compare.');
    }
  } catch (error) {
    alert('Error selecting files: ' + error);
  }
});

// Add event listener for the compare button
document.getElementById('compare-btn').addEventListener('click', async () => {
  if (selectedFiles.length < 2) return;
  
  const loadingEl = document.getElementById('loading');
  const resultsEl = document.getElementById('results');
  
  loadingEl.style.display = 'block';
  resultsEl.style.display = 'none';
  
  try {
    const comparison = await window.api.compareFiles(selectedFiles[0], selectedFiles[1]);
    
    loadingEl.style.display = 'none';
    
    if (comparison.success) {
      displayResults(comparison.results);
      resultsEl.style.display = 'block';
    } else {
      alert('Error comparing files: ' + comparison.error);
    }
  } catch (error) {
    loadingEl.style.display = 'none';
    alert('Error comparing files: ' + error);
  }
});

/**
 * Display the comparison results in the UI
 * @param {Object} results - The comparison results
 */
function displayResults(results) {
  const summaryEl = document.getElementById('summary');
  const differencesEl = document.getElementById('differences');
  
  // Clear previous results
  differencesEl.innerHTML = '';
  
  // Check if there are any differences
  if (results.differences.length === 0) {
    summaryEl.textContent = 'The Excel files are identical! No differences were found.';
    return;
  }
  
  // Display summary
  summaryEl.textContent = `Found ${results.differences.length} differences between the Excel files.`;
  
  // Create table container with resizable functionality
  const tableContainer = document.createElement('div');
  tableContainer.className = 'resizable-table-container';
  
  // Create table for differences
  const table = document.createElement('table');
  table.className = 'diff-table resizable-table';
  
  // Create header row
  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  const headers = ['Sheet', 'Cell', 'File 1 Value', 'File 2 Value'];
  
  headers.forEach((headerText, index) => {
    const header = document.createElement('th');
    header.textContent = headerText;
    header.className = 'resizable';
    
    // Add resizer element except for the last header
    if (index < headers.length - 1) {
      const resizer = document.createElement('div');
      resizer.className = 'resizer';
      header.appendChild(resizer);
      
      // Make column resizable
      resizer.addEventListener('mousedown', function(e) {
        // Get the current width of the column
        const headerWidth = header.offsetWidth;
        const startX = e.pageX;
        
        // Function to handle mouse movement
        function onMouseMove(e) {
          // Calculate the new width
          const newWidth = headerWidth + (e.pageX - startX);
          
          // Update the column width
          if (newWidth > 50) { // Minimum width
            header.style.width = newWidth + 'px';
          }
        }
        
        // Function to handle mouse up
        function onMouseUp() {
          // Remove event listeners
          document.removeEventListener('mousemove', onMouseMove);
          document.removeEventListener('mouseup', onMouseUp);
        }
        
        // Add event listeners
        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);
        
        e.preventDefault(); // Prevent text selection
      });
    }
    
    headerRow.appendChild(header);
  });
  
  thead.appendChild(headerRow);
  table.appendChild(thead);
  
  // Create table body
  const tbody = document.createElement('tbody');
  
  // Create data rows for each difference
  results.differences.forEach((diff, rowIndex) => {
    const row = document.createElement('tr');
    row.className = 'different';
    
    const sheetCell = document.createElement('td');
    sheetCell.textContent = diff.sheet;
    
    const cellIdCell = document.createElement('td');
    cellIdCell.textContent = diff.cell;
    
    const value1Cell = document.createElement('td');
    value1Cell.textContent = diff.value1;
    value1Cell.className = 'resizable-cell';
    
    const value2Cell = document.createElement('td');
    value2Cell.textContent = diff.value2;
    value2Cell.className = 'resizable-cell';
    
    // Add double-click event to expand cell height
    value1Cell.addEventListener('dblclick', function() {
      this.classList.toggle('expanded');
    });
    
    value2Cell.addEventListener('dblclick', function() {
      this.classList.toggle('expanded');
    });
    
    row.appendChild(sheetCell);
    row.appendChild(cellIdCell);
    row.appendChild(value1Cell);
    row.appendChild(value2Cell);
    
    tbody.appendChild(row);
  });
  
  table.appendChild(tbody);
  tableContainer.appendChild(table);
  
  // Add instructions for resizing
  const instructions = document.createElement('p');
  instructions.className = 'resize-instructions';
  instructions.innerHTML = 
    'Drag column dividers to adjust width. Double-click a cell to expand/collapse its height.';
  
  // Add the table to the differences container
  differencesEl.appendChild(instructions);
  differencesEl.appendChild(tableContainer);
}