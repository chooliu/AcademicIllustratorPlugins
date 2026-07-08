// unstretch-text.jsx
// version: 1.0 | 2026-07-07 | Choo Liu
// github.com/chooliu/AcademicIllustratorPlugins
// resets non-uniform character height and weight scaling back to 100%
// (useful after manually reshaping a fig and accidentally scalling text)
// arguments: (none)

(function () {
    if (app.documents.length === 0) { alert("open a document first."); return; }
    var doc = app.activeDocument;

    var frames = (doc.selection && doc.selection.length > 0)
        ? collectTextFrames(doc.selection)
        : toArray(doc.textFrames);

    if (frames.length === 0) { alert("no text objects found."); return; }

    var fixed = 0;
    for (var i = 0; i < frames.length; i++) {
        if (unstretch(frames[i])) fixed++;
    }
    app.redraw();
    alert("unstretched " + fixed + " of " + frames.length + " text object(s).");

    function unstretch(tf) {
        var a = tf.textRange.characterAttributes;
        if (a.horizontalScale === 100 && a.verticalScale === 100) return false;
        a.horizontalScale = 100;
        a.verticalScale = 100;
        return true;
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
