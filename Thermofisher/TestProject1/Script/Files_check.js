function dee() {
  function CountFilesAndTypesInDirectory(directoryPath) {
    var fsObj = Sys.OleObject("Scripting.FileSystemObject");
    var mainFolder = fsObj.GetFolder(directoryPath);

    function processFolder(folder) {
      var fileCount = 0;
      var fileTypes = {};

      var filesEnum = new Enumerator(folder.Files);
      for (; !filesEnum.atEnd(); filesEnum.moveNext()) {
        var file = filesEnum.item();
        var fileName = file.Name;
        var fileType = fileName.split('.').pop().toLowerCase(); // Extract file extension

        if (!fileTypes[fileType]) {
          fileTypes[fileType] = 1;
        } else {
          fileTypes[fileType]++;
        }

        fileCount++;
      }

      // Log the file types and counts for the current folder
      Log.Message("Folder: " + folder.Path);
      Log.Message("Number of files in folder: " + fileCount);

      var fileTypesMessage = "File types and counts:";
      for (var type in fileTypes) {
        fileTypesMessage += " " + type + ": " + fileTypes[type];
      }
      Log.Message(fileTypesMessage);

      var subfoldersEnum = new Enumerator(folder.Subfolders);
      if (!subfoldersEnum.atEnd()) {
        // Log a single line for subfolders with all file types and counts
        var subfolderTypes = [];
        do {
          var subfolder = subfoldersEnum.item();
          var subfileTypes = processFolder(subfolder);
          subfolderTypes.push(subfolder.Name + ": " + subfileTypes.join(", "));
          subfoldersEnum.moveNext();
        } while (!subfoldersEnum.atEnd());

        Log.Message("File types and counts in subfolders: " + subfolderTypes.join(", "));
      }

      return Object.keys(fileTypes).map(function (key) {
        return key + ': ' + fileTypes[key];
      });
    }

    processFolder(mainFolder); // Start processing from the main folder
  }

  // Usage example
  var directoryPath = "C:\\Program Files\\Thermo Scientific\\Ardia";
  CountFilesAndTypesInDirectory(directoryPath);
}
