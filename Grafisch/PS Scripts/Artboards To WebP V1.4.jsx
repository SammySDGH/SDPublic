/*
Artboards-to-WEBP-PS2022.jsx
14th May 2022, v1.1 - Stephen Marsh

* Only for Photoshop 2022 (v23.x.x) or later!
* Saves artboards as lossy WebP - Quality: 100
* Prompts for a save location and saves using the filename + artboard name
* Existing files will be silently overwritten

Instructions for saving and installing:
https://prepression.blogspot.com/2017/11/downloading-and-installing-adobe-scripts.html
*/

// Original script author below:

// =========================================================================
// artboardsToPNG.jsx - Adobe Photoshop Script
// Version: 0.6.0
// Requirements: Adobe Photoshop CC 2015, or higher
// Author: Anton Lyubushkin (nvkz.nemo@gmail.com)
// Website: http://lyubushkin.pro/
// =========================================================================

#target photoshop

app.bringToFront();

// Ensure that version 2022 or later is being used
var versionNumber = app.version.split(".");
var versionCheck = parseInt(versionNumber);
if (versionCheck < 23) {
    alert("You must use Photoshop 2022 or later to save using native WebP format...");

} else {
    if (app.documents.length !== 0) {

        var docRef = app.activeDocument,
            allArtboards,
            artboardsCount = 0,
            inputFolder = Folder.selectDialog("Select an output folder to save the artboards as WebP:");

        if (inputFolder) {
            function getAllArtboards() {
                try {
                    var ab = [];
                    var theRef = new ActionReference();
                    theRef.putProperty(charIDToTypeID('Prpr'), stringIDToTypeID("artboards"));
                    theRef.putEnumerated(charIDToTypeID('Dcmn'), charIDToTypeID('Ordn'), charIDToTypeID('Trgt'));
                    var getDescriptor = new ActionDescriptor();
                    getDescriptor.putReference(stringIDToTypeID("null"), theRef);
                    var abDesc = executeAction(charIDToTypeID("getd"), getDescriptor, DialogModes.NO).getObjectValue(stringIDToTypeID("artboards"));
                    var abCount = abDesc.getList(stringIDToTypeID('list')).count;
                    if (abCount > 0) {
                        for (var i = 0; i < abCount; ++i) {
                            var abObj = abDesc.getList(stringIDToTypeID('list')).getObjectValue(i);
                            var abTopIndex = abObj.getInteger(stringIDToTypeID("top"));
                            ab.push(abTopIndex);

                        }
                    }
                    return [abCount, ab];
                } catch (e) {
                    alert(e.line + '\n' + e.message);
                }
            }

            function selectLayerByIndex(index, add) {
                add = undefined ? add = false : add
                var ref = new ActionReference();
                ref.putIndex(charIDToTypeID("Lyr "), index + 1);
                var desc = new ActionDescriptor();
                desc.putReference(charIDToTypeID("null"), ref);
                if (add) desc.putEnumerated(stringIDToTypeID("selectionModifier"), stringIDToTypeID("selectionModifierType"), stringIDToTypeID("addToSelection"));
                desc.putBoolean(charIDToTypeID("MkVs"), false);
                executeAction(charIDToTypeID("slct"), desc, DialogModes.NO);
            }

            function ungroupLayers() {
                var desc1 = new ActionDescriptor();
                var ref1 = new ActionReference();
                ref1.putEnumerated(charIDToTypeID('Lyr '), charIDToTypeID('Ordn'), charIDToTypeID('Trgt'));
                desc1.putReference(charIDToTypeID('null'), ref1);
                executeAction(stringIDToTypeID('ungroupLayersEvent'), desc1, DialogModes.NO);
            }

            function crop() {
                var desc1 = new ActionDescriptor();
                desc1.putBoolean(charIDToTypeID('Dlt '), true);
                executeAction(charIDToTypeID('Crop'), desc1, DialogModes.NO);
            }

            function saveAsWebP(_name) {
                var s2t = function (s) {
                    return app.stringIDToTypeID(s);
                };
                var descriptor = new ActionDescriptor();
                var descriptor2 = new ActionDescriptor();
                descriptor2.putEnumerated(s2t("compression"), s2t("WebPCompression"), s2t("compressionLossy"));
                descriptor2.putInteger(s2t("quality"), 100);
                descriptor2.putBoolean(s2t("includeXMPData"), true);
                descriptor2.putBoolean(s2t("includeEXIFData"), true);
                descriptor2.putBoolean(s2t("includePsExtras"), true);
                descriptor.putObject(s2t("as"), s2t("WebPFormat"), descriptor2);
                descriptor.putPath(s2t("in"), new File(inputFolder + '/' + _name + '.webp'), true);
                descriptor.putInteger(s2t("documentID"), 237);
                descriptor.putBoolean(s2t("lowerCase"), true);
                descriptor.putEnumerated(s2t("saveStage"), s2t("saveStageType"), s2t("saveSucceeded"));
                executeAction(s2t("save"), descriptor, DialogModes.NO);
            }

            function main(i) {
                selectLayerByIndex(allArtboards[1][i]);
                // RegEx remove filename extension
                var docName = app.activeDocument.name.replace(/\.[^\.]+$/, '');
                // RegEx replace illegal filename characters with a hyphen
                var artboardName = docName + " - " + app.activeDocument.activeLayer.name.replace(/[:\/\\*\?\"\<\>\|\\\r\\\n.]/g, "-"); // "/\:*?"<>|\r\n" -> "-"
                    
                executeAction(stringIDToTypeID("newPlacedLayer"), undefined, DialogModes.NO);
                executeAction(stringIDToTypeID("placedLayerEditContents"), undefined, DialogModes.NO);
                app.activeDocument.selection.selectAll();
                try {
                    ungroupLayers();
                } catch (e) {
                    alert("There was an unexpected error with the ungroupLayers function!")
                    app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);
                }
                crop();
                saveAsWebP(artboardName);
                app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);
            }

            allArtboards = getAllArtboards();

            artboardsCount = allArtboards[0];

            for (var i = 0; i < artboardsCount; i++) {
                docRef.suspendHistory('Artboards to WebP', 'main(' + i + ')');
                app.refresh();
                app.activeDocument.activeHistoryState = app.activeDocument.historyStates[app.activeDocument.historyStates.length - 2];
            }
        }

        app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);
        alert('WebP files saved to: ' + '\r' + inputFolder.fsName);
        // inputFolder.execute();

    } else {
        alert('You must have a document open!');
    }
}