# -*- coding: utf-8 -*-
"""مرتب‌سازی ترتیب فرزندان عناصر XML مطابق ترتیب اسکیمای OOXML.

python-docx عناصر افزوده‌شده را در انتهای والد قرار می‌دهد؛ Word برای برخی
والدها (pPr, rPr, tblPr, sectPr, ...) ترتیب مشخصی انتظار دارد. این ماژول
ترتیب را اصلاح می‌کند تا سند بدون پیام «محتوای خوانا نیست» باز شود.
"""
import re
import shutil
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def q(tag):
    return f"{{{W}}}{tag}"


SEQUENCES = {
    "pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
            "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
            "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE",
            "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
            "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
            "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
            "sectPr", "pPrChange"],
    "rPr": ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
            "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
            "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz",
            "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign",
            "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"],
    "tblPr": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
              "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders",
              "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription"],
    "trPr": ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit",
             "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden", "ins", "del",
             "trPrChange"],
    "tcPr": ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap",
             "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark", "headers",
             "cellIns", "cellDel", "cellMerge", "tcPrChange"],
    "sectPr": ["headerReference", "footerReference", "footnotePr", "endnotePr", "type",
               "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols",
               "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi",
               "rtlGutter", "docGrid", "printerSettings", "sectPrChange"],
    "style": ["name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden",
              "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked", "personal",
              "personalCompose", "personalReply", "rsid", "pPr", "rPr", "tblPr", "trPr",
              "tcPr", "tblStylePr"],
    "pBdr": ["top", "start", "left", "bottom", "end", "right", "between", "bar"],
    "tblBorders": ["top", "start", "left", "bottom", "end", "right", "insideH", "insideV"],
    "tcBorders": ["top", "start", "left", "bottom", "end", "right", "insideH", "insideV",
                  "tl2br", "tr2bl"],
    "tblCellMar": ["top", "start", "left", "bottom", "end", "right"],
    "settings": ["writeProtection", "view", "zoom", "removePersonalInformation",
                 "removeDateAndTime", "doNotDisplayPageBoundaries", "displayBackgroundShape",
                 "printPostScriptOverText", "printFractionalCharacterWidth", "printFormsData",
                 "embedTrueTypeFonts", "embedSystemFonts", "saveSubsetFonts", "saveFormsData",
                 "mirrorMargins", "alignBordersAndEdges", "bordersDoNotSurroundHeader",
                 "bordersDoNotSurroundFooter", "gutterAtTop", "hideSpellingErrors",
                 "hideGrammaticalErrors", "activeWritingStyle", "proofState", "formsDesign",
                 "attachedTemplate", "linkStyles", "stylePaneFormatFilter",
                 "stylePaneSortMethod", "documentType", "mailMerge", "revisionView",
                 "trackChanges", "doNotTrackMoves", "doNotTrackFormatting",
                 "documentProtection", "autoFormatOverride", "styleLockTheme",
                 "styleLockQFSet", "defaultTabStop", "autoHyphenation",
                 "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps",
                 "showEnvelope", "summaryLength", "clickAndTypeStyle", "defaultTableStyle",
                 "evenAndOddHeaders", "bookFoldRevPrinting", "bookFoldPrinting",
                 "bookFoldPrintingSheets", "drawingGridHorizontalSpacing",
                 "drawingGridVerticalSpacing", "displayHorizontalDrawingGridEvery",
                 "displayVerticalDrawingGridEvery", "doNotUseMarginsForDrawingGridOrigin",
                 "drawingGridHorizontalOrigin", "drawingGridVerticalOrigin",
                 "doNotShadeFormData", "noPunctuationKerning", "characterSpacingControl",
                 "printTwoOnOne", "strictFirstAndLastChars", "noLineBreaksAfter",
                 "noLineBreaksBefore", "savePreviewPicture", "doNotValidateAgainstSchema",
                 "saveInvalidXml", "ignoreMixedContent", "alwaysShowPlaceholderText",
                 "doNotDemarcateInvalidXml", "saveXmlDataOnly", "useXSLTWhenSaving",
                 "saveThroughXslt", "showXMLTags", "alwaysMergeEmptyNamespace",
                 "updateFields", "hdrShapeDefaults", "footnotePr", "endnotePr", "compat",
                 "docVars", "rsids", "mathPr", "uiCompat97To2003", "attachedSchema",
                 "themeFontLang", "clrSchemeMapping", "doNotIncludeSubdocsInStats",
                 "doNotAutoCompressPictures", "forceUpgrade", "captions",
                 "readModeInkLockDown", "smartTagType", "schemaLibrary", "shapeDefaults",
                 "doNotEmbedSmartTags", "decimalSymbol", "listSeparator"],
}

PARTS = re.compile(r"word/(document|header\d*|footer\d*|styles|settings|numbering)\.xml$")


def reorder(el, order):
    idx = {name: i for i, name in enumerate(order)}
    children = list(el)
    keyed = []
    for pos, child in enumerate(children):
        local = etree.QName(child).localname
        keyed.append((idx.get(local, len(order) + 1), pos, child))
    keyed.sort(key=lambda t: (t[0], t[1]))
    if [c for _, _, c in keyed] == children:
        return 0
    for child in children:
        el.remove(child)
    for _, _, child in keyed:
        el.append(child)
    return 1


def normalize(path):
    tmp = path + ".tmp"
    changed = 0
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if PARTS.search(item.filename):
                root = etree.fromstring(data)
                for name, order in SEQUENCES.items():
                    for el in root.iter(q(name)):
                        changed += reorder(el, order)
                data = etree.tostring(root, xml_declaration=True,
                                      encoding="UTF-8", standalone=True)
            zout.writestr(item, data)
    shutil.move(tmp, path)
    return changed


if __name__ == "__main__":
    import sys
    for f in sys.argv[1:]:
        print(f, "reordered elements:", normalize(f))
