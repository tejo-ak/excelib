const fs = require('fs');
const vm = require('vm');
// Create a global context to run our scripts in
const globalContext = vm.createContext(global);
// If your script does any DOM manipulation, add the JSDom globals to the globalContext
// globalContext.self = window;
// globalContext.window = window;
// Load the script file and run it in the global context
vm.runInContext(
  fs.readFileSync('dist/excelutil.js'), globalContext,{ filename: 'excelutil.js' }
);