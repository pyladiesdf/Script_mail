function listFilesInFolder(folderName) {

   var sheet = SpreadsheetApp.getActiveSheet();
   sheet.appendRow(["Name", "File-Id"]);


    //Atualizar ID do folder com arquivos
    var folder = DriveApp.getFolderById("FOLDER_ID");
    var contents = folder.getFiles();

    var cnt = 0;
    var file;

    while (contents.hasNext()) {
        var file = contents.next();
        cnt++;

           data = [
                file.getName(),
                file.getId(),
            ];

            sheet.appendRow(data);
    };
};
