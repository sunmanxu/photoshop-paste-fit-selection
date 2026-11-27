#target photoshop

// ============================================================
//  粘贴贴合选区.jsx  ——  Photoshop 脚本
//
//  作用：把剪贴板里的图片粘到当前文档，非等比拉伸、严丝合缝地
//        铺满当前矩形选区，左上角对齐。没选区时退化为普通粘贴。
//
//  怎么用（不需要安装）：
//    1) 在 Photoshop 里画一个矩形选区（不画也行）
//    2) 复制一张图片（网页 / 截图 / 文件夹里的都行）
//    3) 菜单 文件 → 脚本 → 浏览，选中本文件运行
//
//  想常驻菜单（可选）：
//    把本文件复制到
//    C:\Program Files\Adobe\Adobe Photoshop 2025\Presets\Scripts\
//    重启 Photoshop，之后 文件 → 脚本 里会直接出现「粘贴贴合选区」
//
//  兼容：Photoshop 2022 及以上（ExtendScript 通道至今保留），
//        PS 2025 实测可用
// ============================================================

// ===================================================================
// Paste + auto-fit to marquee  (paste+stretch into the current marquee)
//
// How to use:
//   1) draw a rectangular marquee in Photoshop
//   2) copy an image (browser / screenshot / file)
//   3) run this script
// It pastes the clipboard image and stretches it (non-uniform allowed)
// so it exactly fills the marquee, aligned to the marquee's top-left.
// With no marquee it behaves like a normal paste.
// ===================================================================

(function () {

    // ---- 0. a document must be open ----
    var doc = null;
    try { if (app.documents.length > 0) { doc = app.activeDocument; } } catch (e) { doc = null; }
    if (!doc) { alert("\u8BF7\u5148\u6253\u5F00\u4E00\u4E2A\u6587\u6863\u518D\u8FD0\u884C\u8FD9\u4E2A\u811A\u672C\u3002"); return; }

    // ---- 1. read current marquee bounds in px ----
    var left = 0, top = 0, sw = 0, sh = 0, hasSel = false;
    try {
        var b = doc.selection.bounds;          // [left, top, right, bottom]
        left = b[0].as("px");
        top = b[1].as("px");
        sw = b[2].as("px") - left;
        sh = b[3].as("px") - top;
        hasSel = (sw >= 1 && sh >= 1);
    } catch (e) { hasSel = false; }

    // ---- 2. paste clipboard content ----
    var pasted = false;
    var oldDialogs = app.displayDialogs;
    try {
        app.displayDialogs = DialogModes.NO;
        try { app.executeAction(charIDToTypeID("past"), undefined, DialogModes.NO); pasted = true; } catch (e1) { pasted = false; }
        if (!pasted) { try { app.runMenuItem(stringIDToTypeID("paste")); pasted = true; } catch (e2) { pasted = false; } }
    } finally {
        app.displayDialogs = oldDialogs;
    }
    if (!pasted) { alert("\u526A\u8D34\u677F\u91CC\u6CA1\u6709\u53EF\u7C98\u8D34\u7684\u56FE\u7247\uFF0C\u8BF7\u5148\u590D\u5236\u4E00\u5F20\u56FE\u7247\u3002"); return; }

    if (!hasSel) { alert("\u5DF2\u7C98\u8D34\uFF0C\u4F46\u5F53\u524D\u6CA1\u6709\u9009\u533A\uFF0C\u6240\u4EE5\u6309\u539F\u5C3A\u5BF8\u653E\u7F6E\uFF08\u672A\u62C9\u4F38\u586B\u5145\uFF09\u3002"); return; }

    // ---- 3. make the new layer a smart object (non-destructive scaling) ----
    var layer = doc.activeLayer;
    var soOk = false;
    try {
        var isSO = (layer.kind === LayerKind.SMARTOBJECT);
        if (isSO) { soOk = true; }
    } catch (eIs) { }
    if (!soOk) {
        try {
            app.executeAction(stringIDToTypeID("newPlacedLayer"), undefined, DialogModes.NO);
            soOk = true;
        } catch (eSO) { soOk = false; }
    }
    if (!soOk) { alert("\u8F6C\u4E3A\u667A\u80FD\u5BF9\u8C61\u5931\u8D25\uFF0C\u5DF2\u6309\u666E\u901A\u56FE\u5C42\u8D34\u5408\uFF08\u7F29\u653E\u4E0D\u53EF\u56DE\u9000\uFF09\u3002"); }

    // after the action-level conversion the cached layer reference goes stale,
    // so re-acquire it from the document before reading bounds
    layer = doc.activeLayer;

    // ---- 4. stretch + align to the marquee ----
    var oldUnits = app.preferences.rulerUnits;
    try {
        app.preferences.rulerUnits = Units.PIXELS;

        for (var i = 0; i < 4; i++) {
            layer = doc.activeLayer;          // keep the reference fresh each round
            var lb = layer.bounds;
            var lw = lb[2].as("px") - lb[0].as("px");
            var lh = lb[3].as("px") - lb[1].as("px");
            if (lw <= 0 || lh <= 0) { alert("\u7C98\u8D34\u8FDB\u6765\u7684\u56FE\u5C42\u6CA1\u6709\u53EF\u89C1\u5185\u5BB9\uFF0C\u65E0\u6CD5\u8D34\u5408\u9009\u533A\u3002"); return; }

            var needSize = (Math.abs(lw - sw) >= 0.5) || (Math.abs(lh - sh) >= 0.5);
            if (needSize) {
                layer.resize((sw / lw) * 100, (sh / lh) * 100, AnchorPosition.TOPLEFT);
            }
            var nb = layer.bounds;
            var dx = Math.round(left - nb[0].as("px"));
            var dy = Math.round(top - nb[1].as("px"));
            if (dx !== 0 || dy !== 0) { layer.translate(dx, dy); }
            if (!needSize && dx === 0 && dy === 0) { break; }
        }
    } catch (err) {
        alert("\u8D34\u5408\u5931\u8D25" + (err.line ? "\uFF08\u7B2C" + err.line + "\u884C\uFF09" : "") + "\uFF1A" + err.message + "\n\uFF08\u82E5\u56FE\u5C42\u88AB\u9501\u5B9A\uFF0C\u8BF7\u5148\u89E3\u9501\u6216\u6805\u683C\u5316\u540E\u518D\u8BD5\u3002\uFF09");
        return;
    } finally {
        app.preferences.rulerUnits = oldUnits;
    }
})();
