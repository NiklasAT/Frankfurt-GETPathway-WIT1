// === Polyline Nuclear Membrane Intensity Analysis ===
// Version 3.0 - Multi-cell analysis with comprehensive statistics

// === 1. Initial setup and validation ===
if (nImages == 0) {
    exit("Fehler: Kein Bild geöffnet. Bitte öffne zuerst ein Bild.");
}

bildTitel = getTitle();
print("\\Clear"); // Clear log window
print("=== Multi-Cell Nuclear Membran-Intensitäts-Analyse gestartet ===");
print("Bearbeite Bild: " + bildTitel);

// Arrays to store all cell data
cellData = newArray(0);
cellCytoplasmBackgrounds = newArray(0);
cellMembraneAverages = newArray(0);
cellMembraneStdDevs = newArray(0);
cellSegmentMaxima = newArray(0); // Will store all individual segment maxima
cellSegmentCounts = newArray(0); // Track how many segments per cell

cellCount = 0;
totalSegments = 0;

// === 2. Multi-cell measurement loop ===
do {
    cellCount++;
    print("\n=== ZELLE " + cellCount + " ===");
    
    // === 2.1 Cytoplasm background measurement ===
    print("\n--- Schritt 1: Cytoplasma-Hintergrundmessung (Zelle " + cellCount + ") ---");
    run("Select None"); // Clear any previous selection
    setTool("rectangle");
    
    waitForUser("Cytoplasma-Hintergrundmessung - Zelle " + cellCount, 
        "1. Wähle ein Rechteck im Cytoplasma dieser Zelle\n" +
        "2. Vermeide Kern und helle Strukturen\n" +
        "3. Drücke OK wenn die Auswahl stimmt");

    // Check if selection exists
    if (selectionType() == -1) {
        exit("Fehler: Keine Auswahl für Cytoplasma-Messung gefunden.");
    }

    run("Measure");
    if (nResults == 0) {
        exit("Fehler: Cytoplasma-Messung fehlgeschlagen.");
    }

    cytoplasmBackground = getResult("Mean", nResults-1);
    cytoplasmBackgroundStd = getResult("StdDev", nResults-1);
    close("Results");

    // Store cytoplasm data
    cellCytoplasmBackgrounds = Array.concat(cellCytoplasmBackgrounds, cytoplasmBackground);
    
    print("Zelle " + cellCount + " Cytoplasma: " + d2s(cytoplasmBackground, 2) + " ± " + d2s(cytoplasmBackgroundStd, 2));

    // === 2.2 Nuclear membrane line drawing ===
    print("\n--- Schritt 2: Kernmembran-Linie zeichnen (Zelle " + cellCount + ") ---");
    run("Select None"); // Clear previous selection
    setTool("polyline");

    waitForUser("Kernmembran-Linie zeichnen - Zelle " + cellCount, 
        "1. Zeichne eine Segmented Line (Polyline) über die Kernmembran\n" +
        "2. Setze Anker-Punkte entlang der Membran (mindestens 3 Punkte)\n" +
        "3. Doppelklick zum Abschließen\n" +
        "4. Drücke OK wenn die Linie korrekt ist");

    // === 2.3 Selection validation ===
    currentSelectionType = selectionType();
    if (currentSelectionType != 6) { // 6 = polyline in ImageJ
        exit("Fehler: Bitte eine Segmented Line (Polyline) auswählen.\nAktueller Auswahltyp: " + currentSelectionType);
    }

    // Get selection coordinates for validation
    getSelectionCoordinates(x, y);
    nAnker = x.length;

    if (nAnker < 3) {
        exit("Fehler: Die Linie muss mindestens 3 Punkte haben.\nAktuelle Anzahl: " + nAnker);
    }

    anzahlSegmente = nAnker - 1;
    print("Polyline erkannt mit " + nAnker + " Anker-Punkten (" + anzahlSegmente + " Segmente)");

    // === 2.4 Profile analysis ===
    print("\n--- Schritt 3: Intensitätsprofil analysieren (Zelle " + cellCount + ") ---");

    // Make sure we're on the original image and selection is active
    selectWindow(bildTitel);
    if (selectionType() == -1) {
        exit("Fehler: Auswahl wurde verloren. Bitte starte das Makro neu.");
    }

    // Get the profile values
    values = getProfile();
    if (values.length == 0) {
        exit("Fehler: Kein Profil gefunden oder Profil ist leer.");
    }

    print("Profil erstellt mit " + values.length + " Datenpunkten");

    // === 2.5 Background correction ===
    correctedValues = newArray(values.length);
    negativeCount = 0;

    for (i = 0; i < values.length; i++) {
        correctedValues[i] = values[i] - cytoplasmBackground;
        if (correctedValues[i] < 0) {
            correctedValues[i] = 0;
            negativeCount++;
        }
    }

    print("Hintergrund abgezogen (" + d2s(cytoplasmBackground, 2) + ")");
    if (negativeCount > 0) {
        print("Warnung: " + negativeCount + " negative Werte auf 0 gesetzt");
    }

    // === 2.6 Segment analysis ===
    print("\n--- Schritt 4: Segmentanalyse (Zelle " + cellCount + ") ---");

    // Calculate distances along polyline
    getSelectionCoordinates(xCoords, yCoords);
    ankerDistanzen = newArray(nAnker);
    ankerDistanzen[0] = 0;

    totalDistance = 0;
    for (i = 1; i < nAnker; i++) {
        dx = xCoords[i] - xCoords[i-1];
        dy = yCoords[i] - yCoords[i-1];
        segmentDistance = sqrt(dx*dx + dy*dy);
        totalDistance += segmentDistance;
        ankerDistanzen[i] = totalDistance;
    }

    // Convert to profile indices
    profileLength = correctedValues.length;
    ankerIndices = newArray(nAnker);

    for (i = 0; i < nAnker; i++) {
        ankerIndices[i] = Math.round((ankerDistanzen[i] / totalDistance) * (profileLength - 1));
        if (ankerIndices[i] >= profileLength) ankerIndices[i] = profileLength - 1;
    }

    // Analyze segments
    segmentMaxima = newArray(anzahlSegmente);
    segmentMeans = newArray(anzahlSegmente);
    segmentStdDevs = newArray(anzahlSegmente);

    for (seg = 0; seg < anzahlSegmente; seg++) {
        start = ankerIndices[seg];
        end = ankerIndices[seg + 1];
        
        if (start >= end) {
            if (start > 0) start = start - 1;
            if (end < profileLength - 1) end = end + 1;
        }
        
        // Calculate segment statistics
        maxVal = 0;
        sum = 0;
        count = 0;
        
        for (j = start; j < end; j++) {
            if (correctedValues[j] > maxVal) maxVal = correctedValues[j];
            sum += correctedValues[j];
            count++;
        }
        
        if (count == 0) {
            segmentMaxima[seg] = 0;
            segmentMeans[seg] = 0;
            segmentStdDevs[seg] = 0;
            continue;
        }
        
        mean = sum / count;
        
        // Calculate standard deviation
        sumSquares = 0;
        for (j = start; j < end; j++) {
            diff = correctedValues[j] - mean;
            sumSquares += diff * diff;
        }
        stdDev = sqrt(sumSquares / count);
        
        segmentMaxima[seg] = maxVal;
        segmentMeans[seg] = mean;
        segmentStdDevs[seg] = stdDev;
        
        print("Segment " + (seg + 1) + ": Max=" + d2s(maxVal, 2) + ", Mittel=" + d2s(mean, 2) + ", StdDev=" + d2s(stdDev, 2));
    }

    // === 2.7 Calculate cell statistics ===
    cellMembraneAverage = calculateMean(segmentMaxima);
    cellMembraneStdDev = calculateStdDev(segmentMaxima);
    
    // Store data for this cell
    cellMembraneAverages = Array.concat(cellMembraneAverages, cellMembraneAverage);
    cellMembraneStdDevs = Array.concat(cellMembraneStdDevs, cellMembraneStdDev);
    cellSegmentMaxima = Array.concat(cellSegmentMaxima, segmentMaxima);
    cellSegmentCounts = Array.concat(cellSegmentCounts, anzahlSegmente);
    
    totalSegments += anzahlSegmente;
    
    print("Zelle " + cellCount + " Kernmembran-Durchschnitt: " + d2s(cellMembraneAverage, 2) + " ± " + d2s(cellMembraneStdDev, 2));
    
    // === 2.8 Ask user to continue or finish ===
    continueChoice = getBoolean("Weitere Zelle analysieren?", "Weitere Zelle", "Analyse beenden");
    
} while (continueChoice);

// === 3. Final calculations and export ===
print("\n=== FINALE BERECHNUNGEN ===");
print("Gesamtanzahl analysierter Zellen: " + cellCount);
print("Gesamtanzahl Segmente: " + totalSegments);

// Calculate overall statistics
overallMembraneAverage = calculateMean(cellMembraneAverages);
overallMembraneStdDev = calculateStdDev(cellMembraneAverages);
overallCytoplasmAverage = calculateMean(cellCytoplasmBackgrounds);
overallCytoplasmStdDev = calculateStdDev(cellCytoplasmBackgrounds);

// Calculate standard error of the mean
overallMembraneSEM = overallMembraneStdDev / sqrt(cellCount);
overallCytoplasmSEM = overallCytoplasmStdDev / sqrt(cellCount);

print("Gesamt-Kernmembran-Durchschnitt: " + d2s(overallMembraneAverage, 3) + " ± " + d2s(overallMembraneStdDev, 3) + " (SEM: " + d2s(overallMembraneSEM, 3) + ")");
print("Gesamt-Cytoplasma-Durchschnitt: " + d2s(overallCytoplasmAverage, 3) + " ± " + d2s(overallCytoplasmStdDev, 3) + " (SEM: " + d2s(overallCytoplasmSEM, 3) + ")");

// === 4. File export ===
print("\n--- Datei-Export ---");
dir = getDirectory("Wähle Zielordner für CSV-Datei");

// Create safe filename
safeImageName = replaceSpecialChars(bildTitel);
defaultName = safeImageName + "_nuclear_membrane_analysis.csv";
filename = getString("Dateiname für CSV-Datei:", defaultName);

if (filename == "" || filename == "Cancel") {
    exit("Abbruch: Kein Dateiname angegeben.");
}

if (!endsWith(filename, ".csv")) {
    filename = filename + ".csv";
}

path = dir + filename;

// === 5. Create comprehensive results table ===
run("Clear Results");
row = 0;

// === 5.1 Summary statistics ===
setResult("Parameter", row, "Bildname"); setResult("Wert", row, bildTitel); row++;
setResult("Parameter", row, "Anzahl_Zellen"); setResult("Wert", row, cellCount); row++;
setResult("Parameter", row, "Gesamtanzahl_Segmente"); setResult("Wert", row, totalSegments); row++;
setResult("Parameter", row, "Analyse_Zeit"); setResult("Wert", row, d2s(getTime(), 0)); row++;
setResult("Parameter", row, ""); setResult("Wert", row, ""); row++; // Empty row

// Overall membrane statistics
setResult("Parameter", row, "Gesamt_Kernmembran_Mittelwert"); setResult("Wert", row, d2s(overallMembraneAverage, 6)); row++;
setResult("Parameter", row, "Gesamt_Kernmembran_StdDev"); setResult("Wert", row, d2s(overallMembraneStdDev, 6)); row++;
setResult("Parameter", row, "Gesamt_Kernmembran_SEM"); setResult("Wert", row, d2s(overallMembraneSEM, 6)); row++;
setResult("Parameter", row, ""); setResult("Wert", row, ""); row++; // Empty row

// Overall cytoplasm statistics
setResult("Parameter", row, "Gesamt_Cytoplasma_Mittelwert"); setResult("Wert", row, d2s(overallCytoplasmAverage, 6)); row++;
setResult("Parameter", row, "Gesamt_Cytoplasma_StdDev"); setResult("Wert", row, d2s(overallCytoplasmStdDev, 6)); row++;
setResult("Parameter", row, "Gesamt_Cytoplasma_SEM"); setResult("Wert", row, d2s(overallCytoplasmSEM, 6)); row++;
setResult("Parameter", row, ""); setResult("Wert", row, ""); row++; // Empty row

// === 5.2 Individual cell data ===
setResult("Parameter", row, "=== EINZELZELL-DATEN ==="); setResult("Wert", row, ""); row++;

// Headers for cell data
setResult("Zell_ID", row, "Zell_ID");
setResult("Segmentanzahl", row, "Segmentanzahl");
setResult("Kernmembran_Mittelwert", row, "Kernmembran_Mittelwert");
setResult("Kernmembran_StdDev", row, "Kernmembran_StdDev");
setResult("Kernmembran_SEM", row, "Kernmembran_SEM");
setResult("Cytoplasma_Hintergrund", row, "Cytoplasma_Hintergrund");
setResult("Einzelne_Segment_Maxima", row, "Einzelne_Segment_Maxima");
row++;

// Cell data rows
segmentIndex = 0;
for (cell = 0; cell < cellCount; cell++) {
    nSegments = cellSegmentCounts[cell];
    
    // Calculate SEM for this cell
    cellSEM = cellMembraneStdDevs[cell] / sqrt(nSegments);
    
    // Get individual segment maxima for this cell
    segmentMaximaString = "";
    for (seg = 0; seg < nSegments; seg++) {
        if (seg > 0) segmentMaximaString += ";";
        segmentMaximaString += d2s(cellSegmentMaxima[segmentIndex + seg], 3);
    }
    segmentIndex += nSegments;
    
    // Add cell data
    setResult("Zell_ID", row, cell + 1);
    setResult("Segmentanzahl", row, nSegments);
    setResult("Kernmembran_Mittelwert", row, d2s(cellMembraneAverages[cell], 6));
    setResult("Kernmembran_StdDev", row, d2s(cellMembraneStdDevs[cell], 6));
    setResult("Kernmembran_SEM", row, d2s(cellSEM, 6));
    setResult("Cytoplasma_Hintergrund", row, d2s(cellCytoplasmBackgrounds[cell], 6));
    setResult("Einzelne_Segment_Maxima", row, segmentMaximaString);
    row++;
}

updateResults();

// === 6. Save results ===
saveAs("Results", path);
print("✓ Ergebnisse erfolgreich gespeichert: " + path);

// === 7. Final summary ===
print("\n=== ANALYSE VOLLSTÄNDIG ABGESCHLOSSEN ===");
print("Analysierte Zellen: " + cellCount);
print("Analysierte Segmente: " + totalSegments);
print("Durchschnittliche Kernmembran-Intensität: " + d2s(overallMembraneAverage, 3) + " ± " + d2s(overallMembraneStdDev, 3));
print("Durchschnittlicher Cytoplasma-Hintergrund: " + d2s(overallCytoplasmAverage, 3) + " ± " + d2s(overallCytoplasmStdDev, 3));
print("Ergebnisse gespeichert in: " + filename);
print("==========================================");

// === Helper Functions ===
function calculateMean(array) {
    if (array.length == 0) return 0;
    sum = 0;
    for (i = 0; i < array.length; i++) {
        sum += array[i];
    }
    return sum / array.length;
}

function calculateStdDev(array) {
    if (array.length <= 1) return 0;
    mean = calculateMean(array);
    sumSquares = 0;
    for (i = 0; i < array.length; i++) {
        diff = array[i] - mean;
        sumSquares += diff * diff;
    }
    return sqrt(sumSquares / (array.length - 1)); // Sample standard deviation
}

function replaceSpecialChars(text) {
    // Replace special characters with hyphens
    text = replace(text, "\\.", "-");
    text = replace(text, " ", "-");
    text = replace(text, "_", "-");
    text = replace(text, "\\(", "-");
    text = replace(text, "\\)", "-");
    text = replace(text, "\\[", "-");
    text = replace(text, "\\]", "-");
    text = replace(text, "\\{", "-");
    text = replace(text, "\\}", "-");
    text = replace(text, "\\#", "-");
    text = replace(text, "\\%", "-");
    text = replace(text, "\\&", "-");
    text = replace(text, "\\*", "-");
    text = replace(text, "\\+", "-");
    text = replace(text, "\\=", "-");
    text = replace(text, "\\?", "-");
    text = replace(text, "\\!", "-");
    text = replace(text, "\\@", "-");
    text = replace(text, "\\$", "-");
    text = replace(text, "\\^", "-");
    text = replace(text, "\\|", "-");
    text = replace(text, "\\\\", "-");
    text = replace(text, "\\/", "-");
    text = replace(text, "\\:", "-");
    text = replace(text, "\\;", "-");
    text = replace(text, "\\<", "-");
    text = replace(text, "\\>", "-");
    text = replace(text, "\\,", "-");
    text = replace(text, "\\\"", "-");
    text = replace(text, "\\'", "-");
    
    // Remove multiple consecutive hyphens
    while (indexOf(text, "--") >= 0) {
        text = replace(text, "--", "-");
    }
    
    // Remove leading/trailing hyphens
    if (startsWith(text, "-")) text = substring(text, 1);
    if (endsWith(text, "-")) text = substring(text, 0, lengthOf(text) - 1);
    
    return text;
}