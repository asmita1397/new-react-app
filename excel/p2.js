function analyzeExcel(path) {
    var app, worksheets;
    app = Application('Microsoft Excel');
    app.includeStandardAdditions = true;
    app.activate();
    app.open("https://myclassbook.org/wp-content/uploads/2017/12/Test.xlsx");
    worksheets = app.worksheets;
  
    // Create JSON for each worksheet.
    for (var i = 0; i < worksheets.length; i++) {
      makeJSON(worksheets[i]);
    }
  
    function makeJSON(ws) {
      var data = {};
      var title = [];
      var first_row = ws.rows[0];
      var worksheet_name = ws.name();
  
      // Acquire line 1 as heading.
      for (var col_i = 0; ; col_i++) {
         var check = first_row.columns[col_i].value();
         if (!check) {
            break;
         }
         title.push(check);
      }
  
      // Acquire as data from the second line
      for (var row_i = 1; ; row_i++) {
         var row = ws.rows[row_i];
         var row_data = {};
         // Determine the presence or absence of data in the first column.
         var id = row.columns[0].value();
         if (!id) {
            break;
         } else {
            data[id] = {};
         }
         // Combine data row by row.
         // If it has the same heading, put it in an array.
         for (var i = 1; i < title.length; i++) {
            if (row_data[title[i]]) {
               // 配列化
               if (!Array.isArray(row_data[title[i]])) {
                  row_data[title[i]] = [row_data[title[i]]];
               }
               // Add
               if (row.columns[i].value()) {
                  row_data[title[i]].push(row.columns[i].value());
               }
            } else {
               row_data[title[i]] = row.columns[i].value();
            }
         }
         data[id] = row_data;
      }
  
      // Export settings
      var filePath = app.chooseFileName({
         defaultName: worksheet_name + '.json',
         defaultLocation: app.pathTo('desktop')
      });
  
      // Write JSON data (convert character code to UTF-8)
      ObjC.import('Cocoa');
      var text = JSON.stringify(data, null, '  ');
      string = $.NSString.stringWithString(text);
      string.writeToFileAtomicallyEncodingError(
        filePath.toString(),
        true,
        $.NSUTF8StringEncoding,
        $()
      );
    }
  }
  
  // Processing when dragging and dropping a file to an application.
  var SystemEvents = Application("System Events");
  var fileTypesToProcess = ["ELSX"];
  var extensionsToProcess = ["xlsx"];
  var typeIdentifiersToProcess = [];
  function openDocuments(droppedItems) {
    for (var item of droppedItems) {
      var alias = SystemEvents.aliases.byName(item.toString());
      var extension = alias.nameExtension();
      var fileType = alias.fileType();
      var typeIdentifier = alias.typeIdentifier();
      if (
         fileTypesToProcess.includes(fileType) 
         || extensionsToProcess.includes(extension)
         || typeIdentifiersToProcess.includes(typeIdentifier)
      ) {
        var path = Path(item.toString().slice(1));
        analyzeExcel(path);
      }
    }
  }
  
  // Describe how to use when double-clicking the application icon
  function run() {
    var sys = Application("System Events");
    sys.includeStandardAdditions = true;
    sys.displayDialog("Please drag and drop the Excel file (xlsx). Convert to JSON file.");
  }
  