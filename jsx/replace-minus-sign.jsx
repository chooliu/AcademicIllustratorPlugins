// replace-minus-sign.jsx
// version: 1.0 | 2026-07-08 | Choo Liu
// github.com/chooliu/AcademicIllustratorPlugins
// replace unicode minus sign (u+2212) with ascii hyphen (-)
// (common issue via ggplot2 axis labels)
// arguments: (none)

(function () {
    if (app.documents.length === 0) { alert("open a document first."); return; }
    var doc = app.activeDocument;

    var UNICODE_MINUS = String.fromCharCode(8722); // u+2212, ggplot2 default negative-number glyph
    var ASCII_HYPHEN = "-";

    var frames = (doc.selection && doc.selection.length > 0)
        ? collectTextFrames(doc.selection)
        : toArray(doc.textFrames);

    if (frames.length === 0) { alert("no text objects found."); return; }

    var changed = 0;
    for (var i = 0; i < frames.length; i++) {
        try {
            changed += replaceMinus(frames[i]);
        } catch (e) {
            // skip text frames with issues
        }
    }

    app.redraw();
    alert("replaced " + changed + " unicode minus sign(s) with ascii hyphen.");

    function replaceMinus(tf) {
        var count = 0;
        if (!tf || !tf.textRange) return 0;
        var range = tf.textRange;
        if (!range || range.length === 0) return 0;

        // walk each character in the text frame
        for (var i = 0; i < range.length; i++) {
            try {
                var ch = range.characters[i];
                if (ch && ch.contents === UNICODE_MINUS) {
                    ch.contents = ASCII_HYPHEN;
                    count++;
                }
            } catch (e) {
                // skip individual character issues
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
