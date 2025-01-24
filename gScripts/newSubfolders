/**
 * Crea o ottiene una sottocartella in una cartella specificata.
 *
 * @param {string} parentFolderId - ID della cartella principale.
 * @param {string} folderName - Nome della sottocartella da creare o ottenere.
 * @returns {string} ID della sottocartella.
 */


function newSubfolder(parentFolderId, folderName) {
    Logger.log('Verifica ID cartella: ' + parentFolderId);
    try {
        // Estrai l'ID se è un URL
        if (parentFolderId.indexOf('folders/') !== -1) {
            parentFolderId = parentFolderId.split('/folders/')[1];
            Logger.log('Estratto ID della cartella dall\'URL: ' + parentFolderId);
        }
        
        var parentFolder = DriveApp.getFolderById(parentFolderId);
        Logger.log('Cartella trovata: ' + parentFolder.getName());
        var subfolderIterator = parentFolder.getFoldersByName(folderName);
        var subfolderId;
        if (subfolderIterator.hasNext()) {
            subfolderId = subfolderIterator.next().getId();
            Logger.log('Sottocartella esistente trovata, ID: ' + subfolderId);
        } else {
            var newFolder = parentFolder.createFolder(folderName);
            subfolderId = newFolder.getId();
            Logger.log('Nuova sottocartella creata, ID: ' + subfolderId);
        }
        return subfolderId;
    } catch (error) {
        Logger.log('Errore durante la ricerca o creazione della sottocartella: ' + error.message);
        throw error;
    }
}
