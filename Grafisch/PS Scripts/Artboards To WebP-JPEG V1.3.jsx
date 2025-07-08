/*
Artboards-to-JPEG-and-WEBP-PS2022.jsx
18th June 2024, v1.3 - Stephen Marsh

* Only for Photoshop 2022 (v23.x.x) or later!
* Saves artboards as JPEG - Quality: 12 & lossy WebP - Quality: 100
* Prompts for a save location and prompts for a prefix to be added to the artboard name (uses the doc name by default)
* Existing files will be silently overwritten
* Converts non-RGB to sRGB color space, RGB color mode and 8 bpc
* If RGB the original color space is retained, beware if using ProPhoto RGB!
* Added metadata removal for Document Ancestors, Camera Raw Settings and XMP
* Added a "script running" window to provide feedback during the script execution
* Creates a temporary working document

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

    (function () {

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
                    artboardsCount = 0;

                var outputFolder = Folder.selectDialog("Select an output folder to save the artboards as JPEG & WebP:");
                // Test if Cancel button returns null, then do nothing
                if (outputFolder === null) {
                    app.beep();
                    return;
                }

                var preFix = prompt("Add the prefix including separator character to the artboard name:", docRef.name.replace(/\.[^\.]+$/, '') + "_");
                // Test if Cancel button returns null, then do nothing
                if (preFix === null) {
                    app.beep();
                    return;
                }

                // Hide the Photoshop panels
                app.togglePalettes();

                if (outputFolder) {

                    // Dupe to a temp doc
                    app.activeDocument.duplicate(app.activeDocument.name.replace(/\.[^\.]+$/, ''), false);

                    // Clean out unwanted metadata
                    deleteDocumentAncestorsMetadata();
                    removeCRSmeta();

                    // If the doc isn't in RGB mode
                    if (activeDocument.mode !== DocumentMode.RGB)
                        // Convert to sRGB
                        activeDocument.convertProfile("sRGB IEC61966-2.1", Intent.RELATIVECOLORIMETRIC, true, false);
                    // Ensure that the doc mode is RGB (to correctly handle Indexed Color mode)
                    activeDocument.changeMode(ChangeMode.RGB);
                    activeDocument.bitsPerChannel = BitsPerChannelType.EIGHT;

                    allArtboards = getAllArtboards();
                    artboardsCount = allArtboards[0];

                    // Loop over the artboards and run the main function
                    for (var i = 0; i < artboardsCount; i++) {

                        // Script running notification window - courtesy of William Campbell
                        // https://www.marspremedia.com/download?asset=adobe-script-tutorial-11.zip
                        // https://youtu.be/JXPeLi6uPv4?si=Qx0OVNLAOzDrYPB4
                        var working;
                        working = new Window("palette");
                        working.preferredSize = [300, 80];
                        working.add("statictext");
                        working.t = working.add("statictext");
                        working.add("statictext");
                        working.display = function (message) {
                            this.t.text = message || "Script running, please wait...";
                            this.show();
                            app.refresh();
                        };
                        working.display();

                        // Call the main function
                        docRef.suspendHistory('Artboards to JPEG & WebP', 'main(' + i + ')');
                        app.refresh();
                        app.activeDocument.activeHistoryState = app.activeDocument.historyStates[app.activeDocument.historyStates.length - 2];

                        // Ensure Photoshop has focus before closing the running script notification window
                        app.bringToFront();
                        working.close();
                    }
                }

                app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);

                // Restore the Photoshop panels
                app.togglePalettes();

                // End of script notification
                app.beep();
                alert(artboardsCount + ' WebP and ' + artboardsCount + ' JPEG files saved to: ' + '\r' + outputFolder.fsName);

                // Open the output folder in Windows Explorer or the Mac Finder
                // outputFolder.execute();


                ///// FUNCTIONS /////

                // The main working function
                function main(i) {
                    selectLayerByIndex(allArtboards[1][i]);
                    // RegEx replace illegal filename characters with a hyphen
                    var artboardName = preFix + app.activeDocument.activeLayer.name.replace(/[:\/\\*\?\"\<\>\|\\\r\\\n.]/g, "_"); // "/\:*?"<>|\r\n" -> "-"

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
                    saveAsJPEG(artboardName);
                    saveAsWebP(artboardName);
                    app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);
                }

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
                    app.runMenuItem(stringIDToTypeID('ungroupLayersEvent'));
                }

                function crop() {
                    var desc1 = new ActionDescriptor();
                    desc1.putBoolean(charIDToTypeID('Dlt '), true);
                    executeAction(charIDToTypeID('Crop'), desc1, DialogModes.NO);
                }

                function saveAsJPEG(_name) {
                    removeXMP();
                    var jpgOptns = new JPEGSaveOptions();
                    jpgOptns.formatOptions = FormatOptions.STANDARDBASELINE;
                    jpgOptns.embedColorProfile = true;
                    jpgOptns.matte = MatteType.NONE;
                    jpgOptns.quality = 12;
                    activeDocument.saveAs(new File(outputFolder + '/' + _name + '.jpg'), jpgOptns, true, Extension.LOWERCASE);
                }

                function saveAsWebP(_name) {
                    var s2t = function (s) {
                        return app.stringIDToTypeID(s);
                    };
                    var descriptor = new ActionDescriptor();
                    var descriptor2 = new ActionDescriptor();
                    descriptor2.putEnumerated(s2t("compression"), s2t("WebPCompression"), s2t("compressionLossy")); // Lossy compression
                    descriptor2.putInteger(s2t("quality"), 100); // WebP Quality
                    descriptor2.putBoolean(s2t("includeXMPData"), true); // Include XMP metadata
                    descriptor2.putBoolean(s2t("includeEXIFData"), true); // Include EXIF metadata
                    descriptor2.putBoolean(s2t("includePsExtras"), true); // Include Ps Extras metadata
                    descriptor.putObject(s2t("as"), s2t("WebPFormat"), descriptor2);
                    descriptor.putPath(s2t("in"), new File(outputFolder + '/' + _name + '.webp'), true);
                    descriptor.putBoolean(s2t("lowerCase"), true);
                    descriptor.putEnumerated(s2t("saveStage"), s2t("saveStageType"), s2t("saveSucceeded"));
                    executeAction(s2t("save"), descriptor, DialogModes.NO);
                }

                function removeXMP() {
                    //https://community.adobe.com/t5/photoshop/script-to-remove-all-meta-data-from-the-photo/td-p/10400906
                    if (!documents.length) return;
                    if (ExternalObject.AdobeXMPScript === undefined) ExternalObject.AdobeXMPScript = new ExternalObject("lib:AdobeXMPScript");
                    var xmp = new XMPMeta(activeDocument.xmpMetadata.rawData);
                    XMPUtils.removeProperties(xmp, "", "", XMPConst.REMOVE_ALL_PROPERTIES);
                    app.activeDocument.xmpMetadata.rawData = xmp.serialize();
                }

                function removeCRSmeta() {
                    //community.adobe.com/t5/photoshop/remove-crs-metadata/td-p/10306935
                    if (!documents.length) return;
                    if (ExternalObject.AdobeXMPScript === undefined) ExternalObject.AdobeXMPScript = new ExternalObject("lib:AdobeXMPScript");
                    var xmp = new XMPMeta(app.activeDocument.xmpMetadata.rawData);
                    XMPUtils.removeProperties(xmp, XMPConst.NS_CAMERA_RAW, "", XMPConst.REMOVE_ALL_PROPERTIES);
                    app.activeDocument.xmpMetadata.rawData = xmp.serialize();
                }

                function deleteDocumentAncestorsMetadata() {
                    whatApp = String(app.name); //String version of the app name
                    if (whatApp.search("Photoshop") > 0) { //Check for photoshop specifically, or this will cause errors
                        if (ExternalObject.AdobeXMPScript === undefined) ExternalObject.AdobeXMPScript = new ExternalObject("lib:AdobeXMPScript");
                        var xmp = new XMPMeta(activeDocument.xmpMetadata.rawData);
                        // Begone foul Document Ancestors!
                        xmp.deleteProperty(XMPConst.NS_PHOTOSHOP, "DocumentAncestors");
                        app.activeDocument.xmpMetadata.rawData = xmp.serialize();
                    }
                }

            } else {
                alert('You must have a document open!');
            }
        }

    })();