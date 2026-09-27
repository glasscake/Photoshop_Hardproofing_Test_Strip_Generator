#target photoshop

function main() {
    if (!documents.length) {
        alert("Please open an image first before running the test strip script.");
        return;
    }

    // Launch the custom UI
    var config = showSetupDialog();
    
    // If user clicked cancel or closed the window, abort
    if (config === null) return;

    var testMode = config.testMode;
    var stepSize = config.stepSize;
    var negCount = config.negCount;
    var posCount = config.posCount;
    var centerVal = config.centerVal;

    // Validate inputs
    if (isNaN(stepSize) || stepSize <= 0) { alert("Invalid step increment."); return; }
    if (isNaN(negCount) || negCount < 0) { alert("Invalid negative step count."); return; }
    if (isNaN(posCount) || posCount < 0) { alert("Invalid positive step count."); return; }
    if (isNaN(centerVal)) { alert("Invalid center value."); return; }

    // Build asymmetrical array using the starting center value
    var steps = [];
    for (var i = -negCount; i < 0; i++) { steps.push(centerVal + (i * stepSize)); }
    steps.push(centerVal);
    for (var i = 1; i <= posCount; i++) { steps.push(centerVal + (i * stepSize)); }

    var sourceDoc = activeDocument;
    var numSteps = steps.length;

    var srcWidth = sourceDoc.width;
    var srcHeight = sourceDoc.height;
    var res = sourceDoc.resolution;

    // --- PREP THE NON-DESTRUCTIVE STACK ---
    // Duplicate the source document so we don't mess up the user's open file
    var tempDoc = sourceDoc.duplicate("Temp_Do_Not_Save");
    app.activeDocument = tempDoc;
    
    // Convert Background to a standard layer so it can be grouped
    if (tempDoc.layers.length > 0 && tempDoc.layers[tempDoc.layers.length - 1].isBackgroundLayer) {
        tempDoc.layers[tempDoc.layers.length - 1].isBackgroundLayer = false;
    }
    
    // Select all layers and put them in a master group
    selectAllLayersAM();
    groupSelectedLayersAM();
    var stackGroup = tempDoc.activeLayer; // This group now holds the entire non-destructive stack
    
    // THE FIX: Set group to Normal so internal adjustments cannot leak out to other squares!
    stackGroup.blendMode = BlendMode.NORMAL;

    // --- CREATE THE MASTER TEST STRIP CANVAS ---
    var spacing = 50; 
    var totalWidth = (srcWidth * numSteps) + (spacing * (numSteps + 1));
    var totalHeight = srcHeight + (spacing * 2);

    var stripDoc = app.documents.add(totalWidth, totalHeight, res, "Multi-Sample Proof Sheet", NewDocumentMode.RGB, DocumentFill.WHITE);
    
    for (var i = 0; i < numSteps; i++) {
        // Copy the non-destructive stack group into the strip document
        app.activeDocument = tempDoc;
        var dup = stackGroup.duplicate(stripDoc, ElementPlacement.PLACEATBEGINNING);
        
        app.activeDocument = stripDoc;
        stripDoc.activeLayer = dup;
        
        var xOffset = spacing + (i * (srcWidth + spacing));
        var yOffset = spacing;
        
        var bounds = dup.bounds;
        var currentX = bounds[0].value;
        var currentY = bounds[1].value;
        
        dup.translate(xOffset - currentX, yOffset - currentY);
        
        var val = steps[i];
        
        // Skip generating an adjustment layer ONLY if the absolute math equals exactly 0.0
        if (val !== 0.0) {
            // Route to the correct layer generator function
            if (testMode === 1) { applyLightLayerAdjustment("exposure", val); } 
            else if (testMode === 2) { applyLightLayerAdjustment("contrast", Math.round(val)); }
            else if (testMode === 3) { applyLightLayerAdjustment("whites", Math.round(val)); }
            else if (testMode === 4) { applyLightLayerAdjustment("blacks", Math.round(val)); }
            else if (testMode === 5) { applyLightLayerAdjustment("highlights", Math.round(val)); }
            else if (testMode === 6) { applyLightLayerAdjustment("shadows", Math.round(val)); }
            else if (testMode === 7) { applyVibranceLayerAdjustment("vibrance", Math.round(val)); }
            else if (testMode === 8) { applyVibranceLayerAdjustment("saturation", Math.round(val)); }
            else if (testMode === 9) { applyVibranceLayerAdjustment("temperature", Math.round(val)); }
            else if (testMode === 10) { applyVibranceLayerAdjustment("tint", Math.round(val)); }
            else if (testMode === 11) { applyClarityLayerAdjustment("clarity", Math.round(val)); }
            else if (testMode === 12) { applyClarityLayerAdjustment("dehaze", Math.round(val)); }
            else if (testMode === 13) { applyGrainLayerAdjustment(Math.round(val)); }
            else if (testMode === 14) { applyHueSaturationAdjustment("hue", Math.round(val)); }
            else if (testMode === 15) { applyHueSaturationAdjustment("saturation", Math.round(val)); }
            else if (testMode === 16) { applyColorBalanceAdjustment("midtones", "cyanRed", Math.round(val)); }
            else if (testMode === 17) { applyColorBalanceAdjustment("midtones", "magentaGreen", Math.round(val)); }
            else if (testMode === 18) { applyColorBalanceAdjustment("midtones", "yellowBlue", Math.round(val)); }
            else if (testMode === 19) { applyColorBalanceAdjustment("shadows", "cyanRed", Math.round(val)); }
            else if (testMode === 20) { applyColorBalanceAdjustment("shadows", "magentaGreen", Math.round(val)); }
            else if (testMode === 21) { applyColorBalanceAdjustment("shadows", "yellowBlue", Math.round(val)); }
            else if (testMode === 22) { applyColorBalanceAdjustment("highlights", "cyanRed", Math.round(val)); }
            else if (testMode === 23) { applyColorBalanceAdjustment("highlights", "magentaGreen", Math.round(val)); }
            else if (testMode === 24) { applyColorBalanceAdjustment("highlights", "yellowBlue", Math.round(val)); }
            else if (testMode === 25) { applyBlackAndWhiteAdjustment("red", Math.round(val)); }
            else if (testMode === 26) { applyBlackAndWhiteAdjustment("yellow", Math.round(val)); }
            else if (testMode === 27) { applyBlackAndWhiteAdjustment("grain", Math.round(val)); } 
            else if (testMode === 28) { applyBlackAndWhiteAdjustment("cyan", Math.round(val)); }
            else if (testMode === 29) { applyBlackAndWhiteAdjustment("blue", Math.round(val)); }
            else if (testMode === 30) { applyBlackAndWhiteAdjustment("magenta", Math.round(val)); }

            // THE SECOND FIX: Native DOM clipping mask to ensure the new test layer stays locked
            stripDoc.activeLayer.grouped = true;
        }
    }
    
    // Clean up
    tempDoc.close(SaveOptions.DONOTSAVECHANGES);
    app.activeDocument = stripDoc;
    
    var testNames = [
        "(Light) Exposure", "(Light) Contrast", "(Light) Whites", "(Light) Blacks", "(Light) Highlights", "(Light) Shadows",
        "(Color and Vibrance) Vibrance", "(Color and Vibrance) Saturation", "(Color and Vibrance) Temperature", "(Color and Vibrance) Tint", 
        "(Clarity) Clarity", "(Dehaze) Dehaze", "(Grain) Amount", 
        "(Hue/Saturation) Master Hue", "(Hue/Saturation) Master Saturation", 
        "(Color Balance) Midtones: Cyan/Red", "(Color Balance) Midtones: Magenta/Green", "(Color Balance) Midtones: Yellow/Blue",
        "(Color Balance) Shadows: Cyan/Red", "(Color Balance) Shadows: Magenta/Green", "(Color Balance) Shadows: Yellow/Blue",
        "(Color Balance) Highlights: Cyan/Red", "(Color Balance) Highlights: Magenta/Green", "(Color Balance) Highlights: Yellow/Blue",
        "(Black & White) Red", "(Black & White) Yellow", "(Black & White) Green", "(Black & White) Cyan", "(Black & White) Blue", "(Black & White) Magenta"
    ];
    alert(testNames[testMode - 1] + " hard proof strip generated successfully!\nTotal steps: " + numSteps);
}

// --- DOM Helper Functions ---
function selectAllLayersAM() {
    var desc = new ActionDescriptor();
    var ref = new ActionReference();
    ref.putEnumerated( charIDToTypeID( "Lyr " ), charIDToTypeID( "Ordn" ), charIDToTypeID( "Trgt" ) );
    desc.putReference( charIDToTypeID( "null" ), ref );
    executeAction( stringIDToTypeID( "selectAllLayers" ), desc, DialogModes.NO );
}

function groupSelectedLayersAM() {
    var desc = new ActionDescriptor();
    var ref = new ActionReference();
    ref.putClass( stringIDToTypeID( "layerSection" ) );
    desc.putReference( charIDToTypeID( "null" ), ref );
    var ref2 = new ActionReference();
    ref2.putEnumerated( charIDToTypeID( "Lyr " ), charIDToTypeID( "Ordn" ), charIDToTypeID( "Trgt" ) );
    desc.putReference( charIDToTypeID( "From" ), ref2 );
    executeAction( charIDToTypeID( "Mk  " ), desc, DialogModes.NO );
}

function clipToPreviousAM() {
    var desc = new ActionDescriptor();
    var ref = new ActionReference();
    ref.putEnumerated( charIDToTypeID( "Lyr " ), charIDToTypeID( "Ordn" ), charIDToTypeID( "Trgt" ) );
    desc.putReference( charIDToTypeID( "null" ), ref );
    // THE FIX: Changed stringIDToTypeID to charIDToTypeID for "GrpL"
    executeAction( charIDToTypeID( "GrpL" ), desc, DialogModes.NO );
}

// Custom ScriptUI Dialog
function showSetupDialog() {
    var win = new Window("dialog", "Hard Proof Generator Suite");
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];

    // Panel 1: Test Selection
    var pnlType = win.add("panel", undefined, "1. Select Test Type");
    pnlType.alignChildren = ["fill", "top"];
    pnlType.margins = 15;
    
    var typeList = [
        "1 - (Light) Exposure",
        "2 - (Light) Contrast",
        "3 - (Light) Whites",
        "4 - (Light) Blacks",
        "5 - (Light) Highlights",
        "6 - (Light) Shadows",
        "7 - (Color and Vibrance) Vibrance",
        "8 - (Color and Vibrance) Saturation",
        "9 - (Color and Vibrance) Temperature",
        "10 - (Color and Vibrance) Tint",
        "11 - (Clarity) Clarity",
        "12 - (Dehaze) Dehaze",
        "13 - (Grain) Amount",
        "14 - (Hue/Saturation) Master Hue",
        "15 - (Hue/Saturation) Master Saturation",
        "16 - (Color Balance) Midtones: Cyan/Red",
        "17 - (Color Balance) Midtones: Magenta/Green",
        "18 - (Color Balance) Midtones: Yellow/Blue",
        "19 - (Color Balance) Shadows: Cyan/Red",
        "20 - (Color Balance) Shadows: Magenta/Green",
        "21 - (Color Balance) Shadows: Yellow/Blue",
        "22 - (Color Balance) Highlights: Cyan/Red",
        "23 - (Color Balance) Highlights: Magenta/Green",
        "24 - (Color Balance) Highlights: Yellow/Blue",
        "25 - (Black & White) Red",
        "26 - (Black & White) Yellow",
        "27 - (Black & White) Green",
        "28 - (Black & White) Cyan",
        "29 - (Black & White) Blue",
        "30 - (Black & White) Magenta"
    ];
    var ddType = pnlType.add("dropdownlist", undefined, typeList);
    ddType.selection = 0; // Default to Exposure

    // Panel 2: Step Parameters
    var pnlSteps = win.add("panel", undefined, "2. Configure Steps");
    pnlSteps.alignChildren = ["left", "top"];
    pnlSteps.margins = 15;

    // --- NEW: STARTING CENTER VALUE BOX ---
    var grpCenter = pnlSteps.add("group");
    grpCenter.add("statictext", undefined, "Starting Value:");
    var inpCenter = grpCenter.add("edittext", undefined, "0");
    inpCenter.characters = 6;

    var grpStep = pnlSteps.add("group");
    grpStep.add("statictext", undefined, "Step Increment:");
    var inpStep = grpStep.add("edittext", undefined, "0.5");
    inpStep.characters = 6;

    var grpNeg = pnlSteps.add("group");
    grpNeg.add("statictext", undefined, "Negative Steps (-):");
    var inpNeg = grpNeg.add("edittext", undefined, "2");
    inpNeg.characters = 4;

    var grpPos = pnlSteps.add("group");
    grpPos.add("statictext", undefined, "Positive Steps (+):");
    var inpPos = grpPos.add("edittext", undefined, "2");
    inpPos.characters = 4;

    // Dynamically update default step size when dropdown changes
    ddType.onChange = function() {
        var idx = ddType.selection.index;
        if (idx === 0) {
            inpStep.text = "0.5"; // Exposure uses decimals
        } else if (idx >= 1 && idx <= 5) {
            inpStep.text = "5"; // Contrast, Whites, Blacks, Highlights, Shadows
        } else if (idx === 8 || idx === 9) {
            inpStep.text = "5"; // Temp, Tint defaults
        } else {
            inpStep.text = "10"; // Everything else needs larger jumps
        }
    }

    // Buttons
    var grpBtns = win.add("group");
    grpBtns.alignment = ["center", "top"];
    var btnOk = grpBtns.add("button", undefined, "Generate Sheet", {name: "ok"});
    var btnCancel = grpBtns.add("button", undefined, "Cancel", {name: "cancel"});

    // Show window and return values if OK is clicked
    if (win.show() === 1) {
        return {
            testMode: ddType.selection.index + 1,
            centerVal: parseFloat(inpCenter.text), // <--- GRAB THE CENTER VALUE
            stepSize: parseFloat(inpStep.text),
            negCount: parseInt(inpNeg.text, 10),
            posCount: parseInt(inpPos.text, 10)
        };
    } else {
        return null;
    }
}

// 1. The Light Layer Generator 
function applyLightLayerAdjustment(targetParam, val) {
    var desc1 = new ActionDescriptor();
    var ref1 = new ActionReference();
    ref1.putClass( stringIDToTypeID( "adjustmentLayer" ) );
    desc1.putReference( charIDToTypeID( "null" ), ref1 );

    var layerData = new ActionDescriptor();
    layerData.putBoolean( stringIDToTypeID( "useLegacy" ), false );
    layerData.putEnumerated( stringIDToTypeID( "brightnessMode" ), stringIDToTypeID( "brightnessMode" ), stringIDToTypeID( "brightnessModeLight" ) );

    layerData.putDouble( stringIDToTypeID( "brightnessExposure" ), (targetParam === "exposure") ? val : 0.0 );
    layerData.putInteger( stringIDToTypeID( "center" ), (targetParam === "contrast") ? val : 0 ); 
    layerData.putInteger( stringIDToTypeID( "brightnessWhites" ), (targetParam === "whites") ? val : 0 );
    layerData.putInteger( stringIDToTypeID( "brightnessBlacks" ), (targetParam === "blacks") ? val : 0 );
    layerData.putInteger( stringIDToTypeID( "brightnessHighlights" ), (targetParam === "highlights") ? val : 0 );
    layerData.putInteger( stringIDToTypeID( "brightnessShadows" ), (targetParam === "shadows") ? val : 0 );

    var descType = new ActionDescriptor();
    descType.putObject( charIDToTypeID( "Type" ), stringIDToTypeID( "brightnessEvent" ), layerData );
    desc1.putObject( charIDToTypeID( "Usng" ), stringIDToTypeID( "adjustmentLayer" ), descType );

    executeAction( charIDToTypeID( "Mk  " ), desc1, DialogModes.NO );
}

// 2. The Vibrance / Color Layer Generator
function applyVibranceLayerAdjustment(targetParam, targetVal) {
    var desc1 = new ActionDescriptor();
    var ref1 = new ActionReference();
    ref1.putClass( stringIDToTypeID( "adjustmentLayer" ) );
    desc1.putReference( charIDToTypeID( "null" ), ref1 );

    var layerData = new ActionDescriptor();
    layerData.putBoolean( stringIDToTypeID( "useLegacy" ), false );
    
    layerData.putInteger( stringIDToTypeID( "vibrance" ), (targetParam === "vibrance") ? targetVal : 0 );
    layerData.putInteger( stringIDToTypeID( "saturation" ), (targetParam === "saturation") ? targetVal : 0 );
    layerData.putInteger( stringIDToTypeID( "temperature" ), (targetParam === "temperature") ? targetVal : 0 );
    layerData.putInteger( stringIDToTypeID( "tint" ), (targetParam === "tint") ? targetVal : 0 );

    var descType = new ActionDescriptor();
    descType.putObject( charIDToTypeID( "Type" ), stringIDToTypeID( "vibrance" ), layerData );
    desc1.putObject( charIDToTypeID( "Usng" ), stringIDToTypeID( "adjustmentLayer" ), descType );

    executeAction( charIDToTypeID( "Mk  " ), desc1, DialogModes.NO );
}

// 3. The Clarity / Dehaze Layer Generator
function applyClarityLayerAdjustment(targetParam, targetVal) {
    var desc1 = new ActionDescriptor();
    var ref1 = new ActionReference();
    ref1.putClass( stringIDToTypeID( "adjustmentLayer" ) );
    desc1.putReference( charIDToTypeID( "null" ), ref1 );

    var layerData = new ActionDescriptor();
    layerData.putInteger( stringIDToTypeID( "clarity" ), (targetParam === "clarity") ? targetVal : 0 );
    layerData.putInteger( stringIDToTypeID( "dehaze" ), (targetParam === "dehaze") ? targetVal : 0 );

    var descType = new ActionDescriptor();
    descType.putObject( charIDToTypeID( "Type" ), stringIDToTypeID( "clarity" ), layerData );
    desc1.putObject( charIDToTypeID( "Usng" ), stringIDToTypeID( "adjustmentLayer" ), descType );

    executeAction( charIDToTypeID( "Mk  " ), desc1, DialogModes.NO );
}

// 4. The Grain Layer Generator
function applyGrainLayerAdjustment(amountVal) {
    var desc1 = new ActionDescriptor();
    var ref1 = new ActionReference();
    ref1.putClass( stringIDToTypeID( "adjustmentLayer" ) );
    desc1.putReference( charIDToTypeID( "null" ), ref1 );

    var layerData = new ActionDescriptor();
    layerData.putInteger( stringIDToTypeID( "grainAmount" ), amountVal );
    layerData.putInteger( stringIDToTypeID( "grainSize" ), 25 );
    layerData.putInteger( stringIDToTypeID( "grainRoughness" ), 50 );
    layerData.putInteger( stringIDToTypeID( "grainSeed" ), 0 );

    var descType = new ActionDescriptor();
    descType.putObject( charIDToTypeID( "Type" ), stringIDToTypeID( "grainAdjustment" ), layerData );
    desc1.putObject( charIDToTypeID( "Usng" ), stringIDToTypeID( "adjustmentLayer" ), descType );

    executeAction( charIDToTypeID( "Mk  " ), desc1, DialogModes.NO );
}

// 5. The Color Balance Layer Generator
function applyColorBalanceAdjustment(targetRange, targetChannel, val) {
    var desc1 = new ActionDescriptor();
    var ref1 = new ActionReference();
    ref1.putClass( stringIDToTypeID( "adjustmentLayer" ) );
    desc1.putReference( charIDToTypeID( "null" ), ref1 );

    var layerData = new ActionDescriptor();
    layerData.putBoolean( stringIDToTypeID( "preserveLuminosity" ), true );

    var listShadows = new ActionList();
    listShadows.putInteger( (targetRange === "shadows" && targetChannel === "cyanRed") ? val : 0 );
    listShadows.putInteger( (targetRange === "shadows" && targetChannel === "magentaGreen") ? val : 0 );
    listShadows.putInteger( (targetRange === "shadows" && targetChannel === "yellowBlue") ? val : 0 );
    
    var listMidtones = new ActionList();
    listMidtones.putInteger( (targetRange === "midtones" && targetChannel === "cyanRed") ? val : 0 );
    listMidtones.putInteger( (targetRange === "midtones" && targetChannel === "magentaGreen") ? val : 0 );
    listMidtones.putInteger( (targetRange === "midtones" && targetChannel === "yellowBlue") ? val : 0 );
    
    var listHighlights = new ActionList();
    listHighlights.putInteger( (targetRange === "highlights" && targetChannel === "cyanRed") ? val : 0 );
    listHighlights.putInteger( (targetRange === "highlights" && targetChannel === "magentaGreen") ? val : 0 );
    listHighlights.putInteger( (targetRange === "highlights" && targetChannel === "yellowBlue") ? val : 0 );

    layerData.putList( stringIDToTypeID( "shadowLevels" ), listShadows );
    layerData.putList( stringIDToTypeID( "midtoneLevels" ), listMidtones );
    layerData.putList( stringIDToTypeID( "highlightLevels" ), listHighlights );

    var descType = new ActionDescriptor();
    descType.putObject( charIDToTypeID( "Type" ), stringIDToTypeID( "colorBalance" ), layerData );
    desc1.putObject( charIDToTypeID( "Usng" ), stringIDToTypeID( "adjustmentLayer" ), descType );

    executeAction( charIDToTypeID( "Mk  " ), desc1, DialogModes.NO );
}

// 6. The Hue/Saturation Layer Generator (Master)
function applyHueSaturationAdjustment(targetParam, val) {
    var desc1 = new ActionDescriptor();
    var ref1 = new ActionReference();
    ref1.putClass( stringIDToTypeID( "adjustmentLayer" ) );
    desc1.putReference( charIDToTypeID( "null" ), ref1 );

    var layerData = new ActionDescriptor();
    layerData.putEnumerated( stringIDToTypeID("presetKind"), stringIDToTypeID("presetKindType"), stringIDToTypeID("presetKindCustom") );
    layerData.putBoolean( stringIDToTypeID( "colorize" ), false );

    var adjList = new ActionList();
    var adjDesc = new ActionDescriptor();
    
    adjDesc.putInteger( stringIDToTypeID( "hue" ), (targetParam === "hue") ? val : 0 );
    adjDesc.putInteger( stringIDToTypeID( "saturation" ), (targetParam === "saturation") ? val : 0 );
    adjDesc.putInteger( stringIDToTypeID( "lightness" ), 0 );
    
    adjList.putObject( stringIDToTypeID( "hueSatAdjustmentV2" ), adjDesc );
    layerData.putList( stringIDToTypeID( "adjustment" ), adjList );

    var descType = new ActionDescriptor();
    descType.putObject( charIDToTypeID( "Type" ), stringIDToTypeID( "hueSaturation" ), layerData );
    desc1.putObject( charIDToTypeID( "Usng" ), stringIDToTypeID( "adjustmentLayer" ), descType );

    executeAction( charIDToTypeID( "Mk  " ), desc1, DialogModes.NO );
}

// 7. The Black & White Layer Generator
function applyBlackAndWhiteAdjustment(targetParam, val) {
    var desc1 = new ActionDescriptor();
    var ref1 = new ActionReference();
    ref1.putClass( stringIDToTypeID( "adjustmentLayer" ) );
    desc1.putReference( charIDToTypeID( "null" ), ref1 );

    var layerData = new ActionDescriptor();
    layerData.putEnumerated( stringIDToTypeID("presetKind"), stringIDToTypeID("presetKindType"), stringIDToTypeID("presetKindCustom") );

    layerData.putInteger( stringIDToTypeID( "red" ), (targetParam === "red") ? val : 40 );
    layerData.putInteger( stringIDToTypeID( "yellow" ), (targetParam === "yellow") ? val : 60 );
    layerData.putInteger( stringIDToTypeID( "grain" ), (targetParam === "grain") ? val : 40 ); 
    layerData.putInteger( stringIDToTypeID( "cyan" ), (targetParam === "cyan") ? val : 60 );
    layerData.putInteger( stringIDToTypeID( "blue" ), (targetParam === "blue") ? val : 20 );
    layerData.putInteger( stringIDToTypeID( "magenta" ), (targetParam === "magenta") ? val : 80 );
    
    layerData.putBoolean( stringIDToTypeID( "useTint" ), false );
    
    var descColor = new ActionDescriptor();
    descColor.putDouble( stringIDToTypeID( "red" ), 225.0 );
    descColor.putDouble( stringIDToTypeID( "grain" ), 211.0 );
    descColor.putDouble( stringIDToTypeID( "blue" ), 179.0 );
    layerData.putObject( stringIDToTypeID( "tintColor" ), stringIDToTypeID( "RGBColor" ), descColor );

    var descType = new ActionDescriptor();
    descType.putObject( charIDToTypeID( "Type" ), stringIDToTypeID( "blackAndWhite" ), layerData );
    desc1.putObject( charIDToTypeID( "Usng" ), stringIDToTypeID( "adjustmentLayer" ), descType );

    executeAction( charIDToTypeID( "Mk  " ), desc1, DialogModes.NO );
}

main();