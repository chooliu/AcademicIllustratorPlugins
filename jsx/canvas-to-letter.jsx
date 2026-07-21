// canvas-to-letter.jsx
// version: 1.00 | 2026-07-21 | Choo Liu
// github.com/chooliu/AcademicIllustratorPlugins
// saves a copy of the active doc enlarging the active canvas to letter (8.5x11in) for manuscript submission,
// as [current_doc]_letter.pdf, with a custom caption in the bottom-right. does not overwrite the original doc.
// relies on some visible layer flattening (see note below at [**])
// arguments:
// - caption text (default: "Figure X")
// - orientation: portrait* (8.5x11) / landscape (11x8.5)
// - anchor existing artwork at: top-left* / center / bottom-left (artwork is inset from its
//   anchored edge(s); e.g. top-left = add margin at left + top, then paste art)
// - artwork margin from the anchored edge(s), in inches (default: 1*)
// - caption margin from page edge, in inches (default: 1*)
// - caption font (default: Arial; falls back to doc default if the named font isn't installed)
// - caption font size, pt (default: 13*)
// - if artwork doesn't fit letter size, will be prompted: scale to fit* / keep size, overflowing

(function () {
    if (app.documents.length === 0) { alert("open the .ai or .pdf file first."); return; }
    var doc = app.activeDocument;

    var PT_PER_IN = 72;
    var LETTER_W_IN = 8.5;
    var LETTER_H_IN = 11;

    var dlg = new Window("dialog", "canvas to letter + caption");
    dlg.alignChildren = "left";
    dlg.spacing = 6;
    dlg.margins = 14;

    dlg.add("statictext", undefined, "saves a new copy as ..._letter.pdf; your original file is not overwritten.");

    var capGrp = dlg.add("group"); capGrp.alignChildren = "right";
    var capInput = addField(capGrp, "caption text:", "Figure X", 24);

    var orientGrp = dlg.add("group"); orientGrp.alignChildren = "left";
    orientGrp.add("statictext", undefined, "orientation:");
    var orientRow = orientGrp.add("group"); orientRow.spacing = 10;
    var portraitRadio = orientRow.add("radiobutton", undefined, "portrait (8.5x11)");
    var landscapeRadio = orientRow.add("radiobutton", undefined, "landscape (11x8.5)");
    portraitRadio.value = true;

    var anchorPanel = dlg.add("panel", undefined, "anchor artwork at:");
    anchorPanel.alignChildren = "left";
    anchorPanel.spacing = 3;
    anchorPanel.margins = 10;
    var anchorTopLeft = anchorPanel.add("radiobutton", undefined, "top-left");
    var anchorCenter = anchorPanel.add("radiobutton", undefined, "center");
    var anchorBottomLeft = anchorPanel.add("radiobutton", undefined, "bottom-left");
    anchorTopLeft.value = true;

    var marginsGrp = dlg.add("group"); marginsGrp.alignChildren = "right"; marginsGrp.spacing = 6;
    var artMarginInput = addField(marginsGrp, "artwork margin (in):", "1", 4);
    var capMarginInput = addField(marginsGrp, "caption margin (in):", "1", 4);
    var fontFamilyInput = addField(marginsGrp, "caption font:", "Arial", 10);
    var fontInput = addField(marginsGrp, "size (pt):", "13", 4);

    // adds a "label: [input]" pair to an existing dialog row; returns the edittext control
    function addField(parent, label, defaultVal, width) {
        parent.add("statictext", undefined, label);
        var input = parent.add("edittext", undefined, defaultVal);
        input.characters = width;
        return input;
    }

    var btnGrp = dlg.add("group");
    btnGrp.add("button", undefined, "OK", { name: "ok" });
    btnGrp.add("button", undefined, "Cancel", { name: "cancel" });

    if (dlg.show() !== 1) return;

    var captionText = capInput.text;
    var landscape = landscapeRadio.value;
    var anchor = anchorCenter.value ? "center" : (anchorBottomLeft.value ? "bottomleft" : "topleft");
    var artMarginIn = toNum(artMarginInput.text);
    var capMarginIn = toNum(capMarginInput.text);
    var fontSize = toNum(fontInput.text);
    var fontFamily = fontFamilyInput.text;

    if (isNaN(artMarginIn) || artMarginIn < 0) { alert("invalid artwork margin."); return; }
    if (isNaN(capMarginIn) || capMarginIn < 0) { alert("invalid caption margin."); return; }
    if (isNaN(fontSize) || fontSize <= 0) { alert("invalid font size."); return; }

    var artMarginPt = artMarginIn * PT_PER_IN;
    var capMarginPt = capMarginIn * PT_PER_IN;
    var pageWPt = (landscape ? LETTER_H_IN : LETTER_W_IN) * PT_PER_IN;
    var pageHPt = (landscape ? LETTER_W_IN : LETTER_H_IN) * PT_PER_IN;

    try {
        resizeCanvas(doc, pageWPt, pageHPt, anchor, artMarginPt);
        addCaption(doc, captionText, fontFamily, fontSize, capMarginPt);
        var outFile = exportLetterPDF(doc);
        app.redraw();
        if (outFile) {
            alert("saved: " + outFile.fsName);
        }
    } catch (e) {
        if (!e || !e.cancelled) alert("error: " + e);
    }

    function resizeCanvas(doc, newW, newH, anchorMode, marginPt) {
        var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()];
        var rect = ab.artboardRect; // [left, top, right, bottom]
        var curW = rect[2] - rect[0];
        var curH = rect[1] - rect[3];

        // margin is inset at the anchored corner
        // or 2x specified margin for center
        var avail = anchorMode === "center" ? 2 * marginPt : marginPt;
        var availW = newW - avail, availH = newH - avail;

        if (curW > availW || curH > availH) {
            var action = askOversizeAction();
            if (action === "cancel") { throw { cancelled: true }; }
            if (action === "scale") {
                var scale = Math.min(availW / curW, availH / curH);
                rect = scaleToFit(doc, ab, rect, scale);
            }
            // "overflow": keep original size; artwork may extend past the page edge
        }

        var left, top;
        if (anchorMode === "center") {
            left = (rect[0] + rect[2]) / 2 - newW / 2;
            top = (rect[1] + rect[3]) / 2 + newH / 2;
        } else if (anchorMode === "bottomleft") {
            left = rect[0] - marginPt;
            top = (rect[3] - marginPt) + newH;
        } else { // topleft
            left = rect[0] - marginPt;
            top = rect[1] + marginPt;
        }
        ab.artboardRect = [left, top, left + newW, top - newH];
    }

    function askOversizeAction() {
        var d = new Window("dialog", "artwork larger than page");
        d.add("statictext", undefined, "the artwork is larger than the letter page.");
        var g = d.add("group");
        var action = "cancel";
        g.add("button", undefined, "scale to fit").onClick = function () { action = "scale"; d.close(); };
        g.add("button", undefined, "keep size").onClick = function () { action = "overflow"; d.close(); };
        g.add("button", undefined, "cancel").onClick = function () { action = "cancel"; d.close(); };
        d.show();
        return action;
    }

    // [**] groups all visible layers into one object ->
    //     scale uniformly (aspect ratio preserved) -> ungroup -> scale
    // this layer flattening is presumably okay for most submission PDF?
    // send github issue if you find this misbehaves!
    function scaleToFit(doc, ab, rect, scale) {
        var pct = scale * 100;
        var items = [];
        for (var li = 0; li < doc.layers.length; li++) collectItems(doc.layers[li], items);

        if (items.length > 0) {
            var grp = doc.layers[0].groupItems.add();
            for (var m = 0; m < items.length; m++) { items[m].locked = false; items[m].move(grp, ElementPlacement.PLACEATEND); }
            grp.resize(pct, pct, true, true, true, true, pct, Transformation.TOPLEFT);
            while (grp.pageItems.length > 0) grp.pageItems[0].move(grp, ElementPlacement.PLACEBEFORE);
            grp.remove();
        }

        ab.artboardRect = [rect[0], rect[1], rect[0] + (rect[2] - rect[0]) * scale, rect[1] - (rect[1] - rect[3]) * scale];
        return ab.artboardRect;
    }

    function collectItems(layer, out) {
        if (!layer.visible) return;
        layer.locked = false;
        var items = layer.pageItems, n = items.length;
        for (var i = 0; i < n; i++) {
            var pi = items[i];
            if (pi.parent === layer) out.push(pi);
        }
        for (var s = 0; s < layer.layers.length; s++) collectItems(layer.layers[s], out);
    }

    function addCaption(doc, text, fontFamily, fontSize, marginPt) {
        if (!text) return;
        var ar = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; // [left, top, right, bottom]

        var tf = doc.textFrames.add();
        tf.contents = text;
        tf.textRange.characterAttributes.size = fontSize;
        if (fontFamily) {
            try { tf.textRange.characterAttributes.textFont = app.textFonts.getByName(fontFamily); }
            catch (e) { /* reset default font if the named one isn't installed */ }
        }

        // nudge the frame's bottom-right corner to the page's bottom-right, minus the margin
        app.redraw();
        var gb = tf.geometricBounds; // [left, top, right, bottom]
        tf.position = [tf.position[0] + (ar[2] - marginPt) - gb[2], tf.position[1] + (ar[3] + marginPt) - gb[3]];
    }

    function exportLetterPDF(doc) {
        var srcName, folder;
        try {
            var f = doc.fullName; // throws if the document was never saved
            srcName = f.name.replace(/\.[^\.]+$/, "");
            folder = f.path;
        } catch (e) {
            srcName = doc.name.replace(/\.[^\.]+$/, "");
            folder = Folder.myDocuments.fsName;
        }
        var outFile = new File(folder + "/" + srcName + "_letter.pdf");
        if (outFile.exists && !confirm(outFile.name + " already exists. overwrite?")) return null;

        var opts = new PDFSaveOptions();
        opts.compatibility = PDFCompatibility.ACROBAT7;
        opts.preserveEditability = true;
        doc.saveAs(outFile, opts);
        return outFile;
    }

    // reject non-numeric, some errors with blanks or "1abc" -> 1
    function toNum(s) {
        return /^\s*-?\d*\.?\d+\s*$/.test(s) ? parseFloat(s) : NaN;
    }
})();
