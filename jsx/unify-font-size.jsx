// unify-font-size.jsx
// version: 1.0 | 2026-07-07 | Choo Liu
// github.com/chooliu/AcademicIllustratorPlugins
// unify text sizes in a specified point range to a single size
// arguments:
// - font pt range in (default: 0.1 to 5)
// - font pt range out (default: 5)
// - selection or whole document (if selection empty)

(function () {
    if (app.documents.length === 0) { alert("open a document first."); return; }
    var doc = app.activeDocument;

    var dlg = new Window("dialog", "unify font sizes");
    dlg.alignChildren = "left";

    var minGrp = dlg.add("group"); minGrp.alignChildren = "right";
    minGrp.add("statictext", undefined, "min size (pt):");
    var minInput = minGrp.add("edittext", undefined, "0.1");
    minInput.characters = 6;

    var maxGrp = dlg.add("group"); maxGrp.alignChildren = "right";
    maxGrp.add("statictext", undefined, "max size (pt):");
    var maxInput = maxGrp.add("edittext", undefined, "5");
    maxInput.characters = 6;

    var targetGrp = dlg.add("group"); targetGrp.alignChildren = "right";
    targetGrp.add("statictext", undefined, "replace with (pt):");
    var targetInput = targetGrp.add("edittext", undefined, "5");
    targetInput.characters = 6;

    var scopeGrp = dlg.add("group"); scopeGrp.alignChildren = "left";
    scopeGrp.add("statictext", undefined, "apply to:");
    var scopeRadios = scopeGrp.add("group"); scopeRadios.orientation = "column";
    var scopeSelection = scopeRadios.add("radiobutton", undefined, "selection (or doc if empty)");
    scopeSelection.value = true;
    scopeRadios.add("radiobutton", undefined, "entire document");

    var btnGrp = dlg.add("group");
    btnGrp.add("button", undefined, "OK", { name: "ok" });
    btnGrp.add("button", undefined, "Cancel", { name: "cancel" });

    if (dlg.show() !== 1) return;

    var minSize = parseFloat(minInput.text);
    var maxSize = parseFloat(maxInput.text);
    var targetSize = parseFloat(targetInput.text);
    var useSelection = scopeSelection.value;

    if (isNaN(minSize) || isNaN(maxSize) || isNaN(targetSize)) {
        alert("invalid input. enter numeric values.");
        return;
    }
    if (minSize > maxSize) {
        alert("min size must be ≤ max size.");
        return;
    }
    if (targetSize <= 0) {
        alert("target size must be > 0.");
        return;
    }

    var frames = useSelection && doc.selection && doc.selection.length > 0
        ? collectTextFrames(doc.selection)
        : toArray(doc.textFrames);

    if (frames.length === 0) { alert("no text objects found."); return; }

    var changed = 0;
    for (var i = 0; i < frames.length; i++) {
        changed += unifyFontSize(frames[i], minSize, maxSize, targetSize);
    }

    app.redraw();
    alert("changed " + changed + " text run(s) in range [" + minSize + "–" + maxSize + "pt] → " + targetSize + "pt.");

    function unifyFontSize(tf, minSz, maxSz, targetSz) {
        var count = 0;
        var textRange = tf.textRange;
        var len = textRange.length;
        for (var i = 0; i < len; i++) {
            var charSize = textRange.characters[i].fontSize;
            if (charSize >= minSz && charSize <= maxSz) {
                textRange.characters[i].fontSize = targetSz;
                count++;
            }
        }
        return count;
    }

    function collectTextFrames(sel) {
        var out = []; for (var i = 0; i < sel.length; i++) walk(sel[i], out); return out;
    }

    function walk(item, out) {
        if (item.typename === "TextFrame") out.push(item);
        else if (item.typename === "GroupItem")
            for (var j = 0; j < item.pageItems.length; j++) walk(item.pageItems[j], out);
    }

    function toArray(coll) {
        var out = []; for (var i = 0; i < coll.length; i++) out.push(coll[i]); return out;
    }
})();
