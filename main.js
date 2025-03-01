const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');

// Import the comparison logic - adjust the path if needed
let compareExcel;
try {
  compareExcel = require('./index');
} catch (error) {
  console.error('Error loading comparison module:', error.message);
  // Fallback function in case the module can't be loaded
  compareExcel = (file1, file2) => {
    return Promise.resolve({
      differences: [],
      error: 'Comparison module not loaded correctly'
    });
  };
}

// Keep a global reference of the window object
let mainWindow;

function createWindow() {
  // Create the browser window
  mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    }
  });

  // Load the index.html file
  mainWindow.loadFile('index.html');

  // Open DevTools in development
  if (process.env.NODE_ENV === 'development') {
    mainWindow.webContents.openDevTools();
  }

  // Handle window closing
  mainWindow.on('closed', function () {
    mainWindow = null;
  });
}

// Create window when Electron is ready
app.whenReady().then(() => {
  createWindow();
  
  // Handle command line arguments if provided
  const args = process.argv.slice(2);
  if (args.length >= 2) {
    console.log('Command line arguments detected, comparing files:', args[0], args[1]);
    try {
      // Run the comparison directly if files are provided via command line
      const file1 = args[0];
      const file2 = args[1];
      
      if (fs.existsSync(file1) && fs.existsSync(file2)) {
        compareExcel(file1, file2)
          .then(results => {
            console.log('Comparison results:', results);
            // You could display these results in the UI or save to a file
          })
          .catch(error => {
            console.error('Error during comparison:', error);
          });
      } else {
        console.error('One or more of the specified files does not exist');
      }
    } catch (error) {
      console.error('Error processing command line arguments:', error);
    }
  }
});

// Quit when all windows are closed
app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', function () {
  if (mainWindow === null) createWindow();
});

// Handle file selection
ipcMain.handle('select-files', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile', 'multiSelections'],
    filters: [
      { name: 'Excel Files', extensions: ['xlsx', 'xls'] }
    ]
  });
  
  return result.filePaths;
});

// Handle comparison
ipcMain.handle('compare-files', async (event, file1Path, file2Path) => {
  try {
    console.log('Comparing files:', file1Path, file2Path);
    const results = await compareExcel(file1Path, file2Path);
    return { success: true, results };
  } catch (error) {
    console.error('Comparison error:', error);
    return { success: false, error: error.message };
  }
});