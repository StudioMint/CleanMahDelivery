/*
<javascriptresource>
<name>$$$/JavaScripts/cleanMyDelivery/Menu=Clean My Delivery...</name>
</javascriptresource>
*/

#target photoshop
var scriptFolder = (new File($.fileName)).parent; // The location of this script

// VARIABLES

var timeStart = null;
var filesList = [];
var saveFolder;
var savedState;
var settingsPaths, settingsChannels, settingsCrop, settingsGuides, settings8bit, settingsMerge, settingsFlatten, settingsFlattenMerge, settingsJPEGQuality, nameExample, resizeJpeg, suffix, replaceFrom, replaceTo;
var formatKeep, formatNew, formatJPG, formatTIFF, formatPSD;
var openDocuments;

try {
    init();
} catch(e) {
    try {
        if (openDocuments) {
            app.activeDocument.activeHistoryState = savedState;
        } else if (openDocuments != undefined) {
            app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);
        }
    } catch(e) {}
    var timeFull = timeSinceStart(timeStart);
    if (timeFull != null) alert("Error code " + e.number + " (line " + e.line + "):\n" + e + "\n\nTime elapsed " + formatSeconds(timeFull));
}

function createDialog() {

    var w = new Window('dialog',"Clean My Delivery",undefined);
        w.alignChildren = "left";
        w.orientation = "column";

        var static_fileCount = w.add("statictext",[0,0,280,20],"0 files selected");

        var grp_browse = w.add("group");
            grp_browse.orientation = "row";
            grp_browse.alignment = "left";

            var btn_browse1 = grp_browse.add ("button",[0,0,91,20],"&Browse...");
                btn_browse1.shortcutKey = "b";
            var btn_browse2 = grp_browse.add ("button",[0,0,91,20],"Current");
            var btn_browse3 = grp_browse.add ("button",[0,0,91,20],"All");

        var grp_settings = w.add("panel",undefined,"Actions");
            grp_settings.orientation = "column";
            grp_settings.alignment = "left";

            var chk_settings8bit = grp_settings.add("checkbox",[0,0,260,20],"Convert to 8 bit");
            var chk_settingsMerge = grp_settings.add("checkbox",[0,0,260,20],"Merge visible");
            var chk_settingsFlatten = grp_settings.add("checkbox",[0,0,260,20],"Flatten layers");
            var chk_settingsFlattenMerge = grp_settings.add("checkbox",[40,0,260,20],"Merge visible before flatten");
            var chk_settingsCrop = grp_settings.add("checkbox",[0,0,260,20],"Crop canvas (flattened enabled)");
            var chk_settingsPaths = grp_settings.add("checkbox",[0,0,260,20],"Remove paths");
            var chk_settingsChannels = grp_settings.add("checkbox",[0,0,260,20],"Remove channels");
            var chk_settingsGuides = grp_settings.add("checkbox",[0,0,260,20],"Remove guides");
            chk_settingsCrop.enabled = false;
            chk_settings8bit.value = true;
            chk_settingsFlatten.value = true;
            chk_settingsCrop.value = true;
            chk_settingsPaths.value = true;
            chk_settingsChannels.value = true;
            chk_settingsGuides.value = true;

        var grp_saving = w.add("panel",undefined,"Saving");
            grp_saving.orientation = "column";
            grp_saving.alignment = "left";
            grp_saving.alignChildren = "left";

            var btn_save = grp_saving.add ("button",[0,0,260,20],"Select location...");

            // var grp_selectFormat = grp_saving.add("group");
            //     grp_selectFormat.orientation = "row";
            //     grp_selectFormat.alignment = "left";

                // var chk_formatKeep = grp_selectFormat.add("radiobutton",[0,0,120,20],"Keep formats");
                // var chk_formatNew = grp_selectFormat.add("radiobutton",[0,0,120,20],"New formats");

                grp_saving.add("statictext", undefined, "Suffix:");
                var edit_suffix = grp_saving.add("edittext", [0,0,260,20], "_MINT_");

                var grp_QuickSuffix = grp_saving.add("group");
                    grp_QuickSuffix.orientation = "row";
                    grp_QuickSuffix.alignment = "right";
                    grp_QuickSuffix.alignChildren = "right";

                    var btn_K1 = grp_QuickSuffix.add ("button",[0,0,80,20],"K1");
                    var btn_K2 = grp_QuickSuffix.add ("button",[0,0,80,20],"K2");
                    var btn_LEV = grp_QuickSuffix.add ("button",[0,0,80,20],"LEV");

                grp_saving.add("statictext", undefined, "Replace:");
                var grp_Replace = grp_saving.add("group");
                    grp_Replace.orientation = "row";
                    grp_Replace.alignment = "right";
                    grp_Replace.alignChildren = "right";

                    var edit_replaceFrom = grp_Replace.add("edittext", [0,0,125,20], "");
                    var edit_replaceTo = grp_Replace.add("edittext", [0,0,125,20], "");

                var static_nameexample = grp_saving.add("statictext", [0,0,260,20], "DocumentName_MINT_");
                    static_nameexample.enabled = false;

                edit_suffix.onChanging = function() { exampleText() }
                edit_replaceFrom.onChanging = function() { exampleText() }
                edit_replaceTo.onChanging = function() { exampleText() }

                grp_saving.add("panel", [0,0,260,1]);
    
                var grp_format1and2 = grp_saving.add("group");
                    grp_format1and2.orientation = "row";
                    grp_format1and2.alignment = "left";
    
                    var grp_format1 = grp_format1and2.add("group");
                        grp_format1.orientation = "column";
                        grp_format1.alignment = "left";
                        grp_format1.alignment = "top";
    
                        var chk_JPEG = grp_format1.add("checkbox",[0,0,90,20],"JPEG");
                        var chk_TIFF = grp_format1.add("checkbox",[0,0,90,20],"TIFF");
                        var chk_PSD = grp_format1.add("checkbox",[0,0,90,20],"PSD");

                        chk_JPEG.onClick = function() { panel_jpegSettings.enabled = !panel_jpegSettings.enabled; }
    
                    var grp_format2 = grp_format1and2.add("group");
                        grp_format2.orientation = "column";
                        grp_format2.alignment = "left";
                        grp_format1.alignment = "top";
    
                        var panel_jpegSettings = grp_format2.add("panel",undefined,"JPEG Settings");
                            panel_jpegSettings.orientation = "column";
                            panel_jpegSettings.alignment = "left";
                            panel_jpegSettings.alignChildren = "left";
    
                            var grp_JPEGQuality = panel_jpegSettings.add("group");
                                grp_JPEGQuality.orientation = "row";
                                grp_JPEGQuality.alignment = "right";
                                grp_JPEGQuality.alignChildren = "right";
    
                                var sli_JPEGQuality = grp_JPEGQuality.add ("slider",[0,0,100,20]);
                                var txt_JPEGSettings = grp_JPEGQuality.add ("statictext",[0,0,15,20],"12");
    
                                    sli_JPEGQuality.value = 12;
                                    sli_JPEGQuality.minvalue = 0;
                                    sli_JPEGQuality.maxvalue = 12;
                                    sli_JPEGQuality.onChanging = function(){ txt_JPEGSettings.text = parseInt(sli_JPEGQuality.value); }
    
                            panel_jpegSettings.add("statictext", undefined, "Fit inside:");
    
                            var grp_resize = panel_jpegSettings.add("group");
                                grp_resize.orientation = "row";
                                grp_resize.alignment = "left";
                                grp_resize.alignChildren = "left";
    
                                var edit_resizeJpeg = grp_resize.add("edittext", [0,0,60,20], "");
                                grp_resize.add("statictext", undefined, "px");

                            panel_jpegSettings.enabled = false;
    
                        grp_format2.add("statictext",[0,0,20,20],"");
                        grp_format2.add("statictext",[0,0,20,20],"");
            
        var grp_Btn = w.add("group");
            grp_Btn.orientation = "row";
            grp_Btn.alignment = "left";
            
            var btn_Close = grp_Btn.add ("button",[0,0,142,20],"Cancel");
            var btn_OK = grp_Btn.add ("button",[0,0,143,20],"OK");

    if (filesList.length > 0) {
        btn_browse1.enabled = false;
        btn_browse2.enabled = false;
        btn_browse3.enabled = false;
        static_fileCount.text = filesList.length + " files from Bridge";
        nameExample = filesList[0].name.substring(0, filesList[0].name.length - 4);
        exampleText()
    }

    // chk_formatNew.value = true;
    chk_TIFF.value = true;
    // formatRadioClick();
    // function formatRadioClick() {
    //     if (chk_formatKeep.value) {
    //         grp_format2.enabled = false;
    //     } else {
    //         grp_format2.enabled = true;
    //     }
    // }

    // chk_formatKeep.onClick = function() {
    //     formatRadioClick();
    // }
    // chk_formatNew.onClick = function() {
    //     formatRadioClick();
    // }

    // chk_JPEG.onClick = function() {
    //     if (chk_JPEG.value) {
    //         sli_JPEGQuality.enabled = true;
    //     } else {
    //         sli_JPEGQuality.enabled = false;
    //     }
    // }

    chk_settingsCrop.onClick = function() {
        if (chk_settingsCrop.value) {
            chk_settingsFlatten.enabled = true;
            chk_settingsFlatten.text = "Flatten layers";
        } else {
            chk_settingsFlatten.enabled = false;
            chk_settingsFlatten.text = "Flatten layers (cropping disabled)";
        }
    }
    chk_settingsFlatten.onClick = function() {
        if (chk_settingsFlatten.value) {
            chk_settingsFlattenMerge.enabled = true;
            chk_settingsMerge.value = false;
            chk_settingsCrop.value = true;
            chk_settingsCrop.enabled = false;
            chk_settingsCrop.text = "Crop canvas (flattened enabled)";
        } else {
            chk_settingsFlattenMerge.enabled = false;
            chk_settingsCrop.enabled = true;
            chk_settingsCrop.value = true;
            chk_settingsCrop.text = "Crop canvas";
        }
    }
    chk_settingsMerge.onClick = function() {
        if (chk_settingsFlatten.value) {
            if (chk_settingsMerge.value) {
                chk_settingsFlatten.value = false;
                chk_settingsFlattenMerge.enabled = false;
                chk_settingsCrop.enabled = true;
                chk_settingsCrop.text = "Crop canvas";
            } else {
                chk_settingsFlatten.value = true;
                chk_settingsFlattenMerge.enabled = true;
            }
        }
    }

    btn_browse1.onClick = function() {
        var selectedFiles = selectFiles();
        if (selectedFiles != null) {
            filesList = selectedFiles;
            openDocuments = false;
            if (filesList.length > 1) {
                static_fileCount.text = filesList.length + " files selected";
            } else {
                static_fileCount.text = "1 file selected";
            }
        }
        nameExample = filesList[0].name.substring(0, filesList[0].name.length - 4);
        exampleText()
    }
    if (app.documents.length == 0) {
        btn_browse2.enabled = false;
        btn_browse3.enabled = false;
    } else if (app.documents.length == 1) {
        btn_browse3.enabled = false;
    }
    btn_browse2.onClick = function() {
        filesList = [activeDocument];
        openDocuments = true;
        static_fileCount.text = "Current document selected";
        nameExample = activeDocument.name.substring(0, activeDocument.name.length - 4);
        exampleText()
    }
    btn_browse3.onClick = function() {
        filesList = [];
        openDocuments = true;
        for (i = 0; i < app.documents.length; i++) {
            filesList.push(app.documents[i]);
        }
        static_fileCount.text = filesList.length + " open documents selected";
        nameExample = activeDocument.name.substring(0, activeDocument.name.length - 4);
        exampleText()
    }

    btn_K1.onClick = function() {
        edit_suffix.text = "_MINT_K1";
        if (edit_resizeJpeg.text == "") edit_resizeJpeg.text = "2000";
        chk_TIFF.value = false;
        chk_JPEG.value = true;
        panel_jpegSettings.enabled = true;
        exampleText()
    }
    btn_K2.onClick = function() {
        edit_suffix.text = "_MINT_K2";
        if (edit_resizeJpeg.text == "") edit_resizeJpeg.text = "2000";
        chk_TIFF.value = false;
        chk_JPEG.value = true;
        panel_jpegSettings.enabled = true;
        exampleText()
    }
    btn_LEV.onClick = function() {
        edit_suffix.text = "_MINT_LEV";
        chk_TIFF.value = true;
        chk_JPEG.value = false;
        panel_jpegSettings.enabled = false;
        exampleText()
    }

    function exampleText() {
        if (!nameExample) {
            static_nameexample.text = replaceWithVar("DocumentName", edit_replaceFrom.text, edit_replaceTo.text) + edit_suffix.text;
        } else {
            static_nameexample.text = replaceWithVar(nameExample, edit_replaceFrom.text, edit_replaceTo.text) + edit_suffix.text;
        }
    }

    btn_save.onClick = function() {
        var selectedFolder = selectFolder();
        if (selectedFolder != null) {
            saveFolder = selectedFolder;
            if (cleanFolderStr(saveFolder).length > 30) {
                var shortSaveFolder = cleanFolderStr(saveFolder);
                    shortSaveFolder = shortSaveFolder.substring(shortSaveFolder.length - 30, shortSaveFolder.length);
                btn_save.text = "..." + shortSaveFolder;
            } else {
                btn_save.text = cleanFolderStr(saveFolder);
            }
        }
    }
    
    btn_OK.onClick = function() {

        if (filesList.length == 0) return alert("You need to select a source");

        if (!chk_settingsCrop.value && !chk_settings8bit.value && !chk_settingsFlatten.value && !chk_settingsPaths.value && !chk_settingsChannels.value && !chk_settingsGuides.value) {
            return alert("No actions selected");
        }

        if (!saveFolder) return alert("Please select a save destination")
        
        if (filesList.length == 0 && app.documents.length == 0) return alert("No files selected");

        if (checkSettings(true) === true) w.close();

    }

    x = w.show();

    // Global variables
    settingsCrop = chk_settingsCrop.value;
    settings8bit = chk_settings8bit.value;
    settingsMerge = chk_settingsMerge.value;
    settingsFlatten = chk_settingsFlatten.value;
    settingsFlattenMerge = chk_settingsFlattenMerge.value;
    settingsPaths = chk_settingsPaths.value;
    settingsChannels = chk_settingsChannels.value;
    settingsGuides = chk_settingsGuides.value;
    settingsJPEGQuality = Number(txt_JPEGSettings.text);

    resizeJpeg = edit_resizeJpeg.text;
    suffix = edit_suffix.text;
    replaceFrom = edit_replaceFrom.text;
    replaceTo = edit_replaceTo.text;

    // formatKeep = chk_formatKeep.value;
    // formatNew = chk_formatNew.value;
    formatJPG = chk_JPEG.value;
    formatTIFF = chk_TIFF.value;
    formatPSD = chk_PSD.value;

    return x;

    function checkSettings(ignoreAlerts) {
        
        // If something has to be checked before closing the window, do it here and return false if faulty

        return true;

    }

}

function init() {

    filesList = getFilesFromBridge();

    // Show window
    var ok = createDialog();
    if (ok === 2) return false;

    // Keeping the ruler settings to reset in the end of the script
    var startRulerUnits = app.preferences.rulerUnits;
    var startTypeUnits = app.preferences.typeUnits;
    var startDisplayDialogs = app.displayDialogs;
    
    // Changing ruler settings to pixels for correct image resizing
    app.preferences.rulerUnits = Units.PIXELS;
    app.preferences.typeUnits = TypeUnits.PIXELS;
    app.displayDialogs = DialogModes.NO;

    // Timer prep
    var d = new Date();
    timeStart = d.getTime() / 1000;

    //// MAIN FUNCTION RUN ////
    for (i = 0; i < filesList.length; i++) {
        var thisItem = filesList[i];
        // currentlyOpen = false;
        // for (j = 0; j < app.documents.length; j++) if (String(thisItem) == String(app.documents[j].fullName)) currentlyOpen = true;
        if (openDocuments) {
            activeDocument = thisItem;
        } else {
            open(File(thisItem));
        }
        activeDocument.suspendHistory("Clean My Delivery", "main()");
    }
    ///////////////////////////

    // Timer calculate
    var timeFull = timeSinceStart(timeStart);

    if (timeFull < 1) {
        alert("Time elapsed: Lightning speed!");
    } else {
        alert("Time elapsed: " + formatSeconds(timeFull));
    }

    // Reset the ruler
    app.preferences.rulerUnits = startRulerUnits;
    app.preferences.typeUnits = startTypeUnits;
    app.displayDialogs = startDisplayDialogs;

}

function main() {

    if (openDocuments) {
        savedState = app.activeDocument.activeHistoryState;
    }
    
    if (settingsPaths) activeDocument.pathItems.removeAll();
    if (settingsChannels) activeDocument.channels.removeAll();
    if (settingsCrop) activeDocument.crop([0, 0, activeDocument.width.value, activeDocument.height.value]);
    if (settingsGuides) {
        var idclearAllGuides = stringIDToTypeID( "clearAllGuides" );
        executeAction( idclearAllGuides, undefined, DialogModes.NO );
    }
    if (settingsFlatten) {
        if (settingsFlattenMerge && activeDocument.layers.length != 1) {
            activeDocument.mergeVisibleLayers(); // Merge all first, because of bug when only flattening
        }
        activeDocument.flatten();
    } else if (settingsMerge && activeDocument.layers.length != 1) {
        activeDocument.mergeVisibleLayers();
    }
    if (settings8bit) activeDocument.bitsPerChannel = BitsPerChannelType.EIGHT;
    
    if (isActiveDocumentSaved()) {
        var documentName = activeDocument.name.substring(0, activeDocument.name.lastIndexOf("."));
    } else {
        var documentName = activeDocument.name;
    }

    documentName = replaceWithVar(documentName, replaceFrom, replaceTo);
    documentName = documentName + suffix;

    if (formatTIFF) saveAsTIFF(saveFolder, documentName);
    if (formatPSD) saveAsPSD(saveFolder, documentName);
    if (formatJPG) {
        if (resizeJpeg != "0" && resizeJpeg != "") {
            if (activeDocument.width.value > activeDocument.height.value) {
                activeDocument.resizeImage(Number(resizeJpeg), null, null, ResampleMethod.BICUBICSHARPER);
            } else {
                activeDocument.resizeImage(null, Number(resizeJpeg), null, ResampleMethod.BICUBICSHARPER);
            }
        }
        saveAsJPG(saveFolder, documentName, settingsJPEGQuality);
    }

    if (openDocuments) {
        app.activeDocument.activeHistoryState = savedState;
    } else {
        app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);
    }
    savedState = undefined;

}

// FUNCTIONS

function saveAsPSD(folder, name) {
    psdFile = new File(folder + "/" + name + ".psd");
    psdSaveOptions = new PhotoshopSaveOptions();
    psdSaveOptions.embedColorProfile = true;
    psdSaveOptions.alphaChannels = true;
    activeDocument.saveAs(psdFile, psdSaveOptions, true, Extension.LOWERCASE);
}

function saveAsJPG(folder, name, quality) {
    jpgFile = new File(folder + "/" + name + ".jpg");
    jpgSaveOptions = new JPEGSaveOptions();
    jpgSaveOptions.embedColorProfile = true;
    jpgSaveOptions.formatOptions = FormatOptions.STANDARDBASELINE;
    jpgSaveOptions.matte = MatteType.NONE;
    jpgSaveOptions.quality = quality;
    activeDocument.saveAs(jpgFile, jpgSaveOptions, true, Extension.LOWERCASE);
}

function saveAsTIFF(folder, name) {
    tiffFile = new File(folder + "/" + name + ".tif");
    tiffSaveOptions = new TiffSaveOptions();
    tiffSaveOptions.embedColorProfile = true;
    tiffSaveOptions.layers = true;
    tiffSaveOptions.imageCompression = TIFFEncoding.TIFFLZW;
    activeDocument.saveAs(tiffFile, tiffSaveOptions, true, Extension.LOWERCASE);
}

function getFilesFromBridge() {

    var fileList = [];
    if (BridgeTalk.isRunning("bridge") && BridgeTalk.getStatus() == "PUMPING") {
        var bt = new BridgeTalk();
        bt.target = "bridge";
        bt.body = "var theFiles = photoshop.getBridgeFileListForAutomateCommand();theFiles.toSource();";
        bt.onResult = function(inBT) { fileList = eval(inBT.body); }
        bt.onError = function(inBT) { fileList = new Array(); }
        bt.send(8);
        return fileList; 
    } else {
        return [];
    }

}

function replaceWithVar(str,replaceWhat,replaceTo){
    if (replaceWhat) {
        var re = new RegExp(replaceWhat, 'g');
        return str.replace(re,replaceTo);
    } else {
        return str;
    }
}

function timeSinceStart(start) {
    if (start == null) return null;
    var d = new Date();
    var timeNow = d.getTime() / 1000;
    return timeNow - start;
}

function formatSeconds(sec) {
    String.prototype.repeat = function(x) {
        var str = "";
        for (var repeats = 0; repeats < x; repeats++) str = str + this;
        return str;
    };
    Number.prototype.twoDigits = function() {
        if (this == 0) return ('0'.repeat(2));
        var dec = this / (Math.pow(10, 2));
        if (String(dec).substring(String(dec).lastIndexOf(".") + 1, String(dec).length).length == 1) dec = dec + "0";
        var str = dec.toString().substring(2, dec.toString().length);
        return str;
    };
    var hours = Math.floor(sec / 60 / 60);
    var minutes = Math.floor(sec / 60) - (hours * 60);
    var seconds = sec % 60;
    return Math.floor(hours).twoDigits() + ':' + Math.floor(minutes).twoDigits() + ':' + Math.floor(seconds).twoDigits();
}

function isActiveDocumentSaved(){
    var ref = new ActionReference();
    ref.putEnumerated( charIDToTypeID('Dcmn'), charIDToTypeID('Ordn'), charIDToTypeID('Trgt') );
    return executeActionGet(ref).hasKey(stringIDToTypeID('fileReference'));
};

function cleanFolderStr(input) { // Input a path in any form and return a string that looks like it's pasted
	
	input = String(input);
	
	if (input[0] == "\/" && isLetter(input[1]) && input[2] == "\/") {
		inputLetter = input[1].toUpperCase();
		input = input.substring(3,input.length);
		input = inputLetter+":\\"+input;
	}
	
	input = input.replace(/\//g, "\\"); // Replace all slashes with backslashes to match if pasted
	input = input.replace(/%20/g, " "); // Replace all %20 with actual spaces to match if pasted
	input = input.replace(/%C3%A9/g, "é"); // Replace all %C3 with é to match if pasted
	
	return input;
	
	function isLetter(str) {
		return str.length === 1 && str.match(/[a-z]/i);
	}
	
}

function selectFiles() {
    var refFiles = File(scriptFolder).openDlg("Select a file", ["*.psd", "*.psb", "*.tif", "*.tiff", "*.jpg", "*.jpeg"], true);
    if (refFiles != null) return refFiles;
}

function selectFolder() {
    var refFolder = Folder.selectDialog("Select a folder");
    if (refFolder != null) return refFolder;
}