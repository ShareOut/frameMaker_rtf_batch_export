// Prompt the folder
var folder = Folder.selectDialog ("Select the folder with the files to export");
var files = folder.getFiles("*.fm");


for (var i = 0; i < files.length; i++) {
    var file = files[i];
    var openParams = GetOpenDefaultParams();

    // Set the params for opening the files
    var j = GetPropIndex(openParams, Constants.FS_FileIsOldVersion);
    openParams[j].propVal.ival = Constants.FV_DoOK;

    var j = GetPropIndex(openParams, Constants.FS_FontNotFoundInCatalog);
    openParams[j].propVal.ival = Constants.FV_DoOK;

    var j = GetPropIndex(openParams, Constants.FS_FontNotFoundInDoc);
    openParams[j].propVal.ival = Constants.FV_DoOK;
    
    var openReturnParams =  new PropVals();
    var doc = Open(file.fsName, openParams, openReturnParams);
    
    saveAsRtf(doc);
    doc.Close(1);
 }


function saveAsRtf(doc) {

    //Make sure we have a valid document object
    if(!doc.ObjectValid()) {
        alert("Document object is not valid. Cannot save to text file.");
        PrintOpenStatus();
        return;
    }

    //Get the default properties for the save
    var props = GetSaveDefaultParams();
    var propIndex;

    //Set the destination to text file
    propIndex = GetPropIndex(props, Constants.FS_FileType);
    props[propIndex].propVal.ival = Constants.FV_SaveFmtText;

    //Set the save mode to "Save As"
    propIndex = GetPropIndex(props, Constants.FS_SaveMode);
    props[propIndex].propVal.ival = Constants.FV_ModeSaveAs;
    
    //Get the current filepath of the document and replace the .fm extension with .rtf
    var filepath = doc.Name;
    filepath = filepath.replace(".fm", ".rtf");

    var returnProps = new PropVals();

    doc.Save(filepath, props, returnProps);
}    