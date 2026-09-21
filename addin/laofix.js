/* Lao Text Fixer — OOXML transform engine
 * phetsarath.mts.la — SIL OFL / free to use
 *
 * Works on the flat OPC package returned by Word JS getOoxml(), or on a
 * bare word/document.xml. Two independent fixes:
 *   1. tagRuns  — mark every run containing Lao text as complex script
 *                 (<w:cs/>, w:szCs, w:rFonts w:cs, w:lang w:bidi="lo-LA").
 *                 This is what makes Word wrap Lao at word boundaries.
 *   2. zwsp     — insert invisible zero-width spaces (U+200B) at Lao
 *                 syllable boundaries, for old Word versions whose layout
 *                 engine still refuses to wrap. Removable with removeZwsp.
 */
(function (root) {
  "use strict";

  var W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  var PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";
  var ZWSP = 0x200b;

  /* CT_RPr child element order (ECMA-376) — inserts must respect it or
     Word rejects the OOXML on insertOoxml(). */
  var RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps",
    "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss",
    "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color",
    "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u",
    "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em",
    "lang", "eastAsianLayout", "specVanish", "oMath"];

  // ---- Lao character classes (mirror of LaoWrapMacro.bas) ----
  function isLao(c) { return c >= 0x0e80 && c <= 0x0eff; }
  function isPreVowel(c) { return c >= 0x0ec0 && c <= 0x0ec4; }
  function isConsonant(c) {
    return (c >= 0x0e81 && c <= 0x0eae) || c === 0x0edc || c === 0x0edd;
  }
  function isVowelSign(c) {
    return (c >= 0x0eb0 && c <= 0x0ebd) || (c >= 0x0ec6 && c <= 0x0ecd);
  }
  function isClusterSecond(c) {
    return c === 0x0ebc || c === 0x0ea7 || c === 0x0ea5 || c === 0x0ea3 ||
           c === 0x0e8d || c === 0x0e99 || c === 0x0ea1 || c === 0x0e87;
  }
  var HAS_LAO = /[຀-໿]/;

  /* Given a paragraph's full text, return the set of character indexes
     that should get a ZWSP inserted BEFORE them. */
  function breakIndexes(s) {
    var out = [];
    for (var i = 1; i < s.length; i++) {
      var prev = s.charCodeAt(i - 1);
      var cur = s.charCodeAt(i);
      var nxt = i + 1 < s.length ? s.charCodeAt(i + 1) : 0;
      if (!isLao(prev) || prev === ZWSP) continue;
      if (isPreVowel(cur)) {
        out.push(i);
      } else if (isConsonant(cur) && !isPreVowel(prev) &&
                 (isVowelSign(nxt) || isPreVowel(nxt) ||
                  (cur === 0x0eab && isClusterSecond(nxt)))) {
        /* never break after a pre-vowel (it must stay glued to its
           consonant), nor inside an initial cluster */
        if (!(isClusterSecond(cur) && isConsonant(prev))) out.push(i);
      }
    }
    return out;
  }

  // ---- DOM helpers ----
  function childNS(el, local) {
    for (var n = el.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 1 && n.localName === local && n.namespaceURI === W) return n;
    }
    return null;
  }

  function ensureRPrChild(doc, rPr, local) {
    var found = childNS(rPr, local);
    if (found) return found;
    var el = doc.createElementNS(W, "w:" + local);
    var idx = RPR_ORDER.indexOf(local);
    var before = null;
    for (var n = rPr.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1 || n.namespaceURI !== W) continue;
      var j = RPR_ORDER.indexOf(n.localName);
      if (j > idx) { before = n; break; }
    }
    rPr.insertBefore(el, before);
    return el;
  }

  function getAttrW(el, local) {
    return el.getAttributeNS(W, local) || el.getAttribute("w:" + local) || "";
  }
  function setAttrW(el, local, val) {
    el.setAttributeNS(W, "w:" + local, val);
  }

  function runText(run) {
    var t = "", ts = run.getElementsByTagNameNS(W, "t");
    for (var i = 0; i < ts.length; i++) t += ts[i].textContent;
    return t;
  }

  /* Scope: the document part only (never styles/numbering parts). */
  function documentRoot(doc) {
    var parts = doc.getElementsByTagNameNS(PKG, "part");
    for (var i = 0; i < parts.length; i++) {
      var name = parts[i].getAttributeNS(PKG, "name") ||
                 parts[i].getAttribute("pkg:name");
      if (name === "/word/document.xml") return parts[i];
    }
    return doc.documentElement;
  }

  // ---- fix 1: complex-script tagging ----
  function tagRun(doc, run, csFont) {
    if (!HAS_LAO.test(runText(run))) return false;
    var rPr = childNS(run, "rPr");
    if (!rPr) {
      rPr = doc.createElementNS(W, "w:rPr");
      run.insertBefore(rPr, run.firstChild);
    }
    var rFonts = ensureRPrChild(doc, rPr, "rFonts");
    if (!getAttrW(rFonts, "cs")) {
      setAttrW(rFonts, "cs",
        getAttrW(rFonts, "ascii") || getAttrW(rFonts, "hAnsi") || csFont);
    }
    ensureRPrChild(doc, rPr, "cs");
    var sz = childNS(rPr, "sz");
    if (sz && !childNS(rPr, "szCs")) {
      setAttrW(ensureRPrChild(doc, rPr, "szCs"), "val", getAttrW(sz, "val"));
    }
    if (childNS(rPr, "b")) ensureRPrChild(doc, rPr, "bCs");
    if (childNS(rPr, "i")) ensureRPrChild(doc, rPr, "iCs");
    var lang = ensureRPrChild(doc, rPr, "lang");
    setAttrW(lang, "bidi", "lo-LA");
    return true;
  }

  // ---- fix 2: ZWSP insertion, paragraph-aware across split runs ----
  function zwspParagraph(para) {
    var ts = para.getElementsByTagNameNS(W, "t");
    if (!ts.length) return 0;
    var full = "", spans = [];
    for (var i = 0; i < ts.length; i++) {
      spans.push({ node: ts[i], start: full.length, len: ts[i].textContent.length });
      full += ts[i].textContent;
    }
    if (!HAS_LAO.test(full)) return 0;
    var idxs = breakIndexes(full);
    if (!idxs.length) return 0;
    /* each break index belongs to exactly one span: the one containing
       the character it precedes; inserting at position 0 of a text node
       is fine, so no cross-span bookkeeping is needed */
    for (var s = 0; s < spans.length; s++) {
      var sp = spans[s], text = sp.node.textContent;
      for (var k = idxs.length - 1; k >= 0; k--) {
        var pos = idxs[k] - sp.start;
        if (pos >= 0 && pos < sp.len) {
          text = text.slice(0, pos) + "\u200B" + text.slice(pos);
        }
      }
      sp.node.textContent = text;
    }
    return idxs.length;
  }

  // ---- public API ----
  function transform(xml, opts, DOMParserImpl, XMLSerializerImpl) {
    opts = opts || {};
    var DP = DOMParserImpl || root.DOMParser;
    var XS = XMLSerializerImpl || root.XMLSerializer;
    var doc = new DP().parseFromString(xml, "text/xml");
    var scope = documentRoot(doc);
    var stats = { runsTagged: 0, zwspInserted: 0, zwspRemoved: 0 };

    if (opts.removeZwsp) {
      var all = scope.getElementsByTagNameNS(W, "t");
      for (var i = 0; i < all.length; i++) {
        var txt = all[i].textContent;
        if (txt.indexOf("\u200B") >= 0) {
          stats.zwspRemoved += txt.split("\u200B").length - 1;
          all[i].textContent = txt.split("\u200B").join("");
        }
      }
    } else {
      if (opts.tagRuns !== false) {
        var runs = scope.getElementsByTagNameNS(W, "r");
        for (var r = 0; r < runs.length; r++) {
          if (tagRun(doc, runs[r], opts.csFont || "Phetsarath OT")) stats.runsTagged++;
        }
      }
      if (opts.zwsp) {
        var paras = scope.getElementsByTagNameNS(W, "p");
        for (var p = 0; p < paras.length; p++) stats.zwspInserted += zwspParagraph(paras[p]);
      }
    }
    return { xml: new XS().serializeToString(doc), stats: stats };
  }

  var api = { transform: transform, breakIndexes: breakIndexes };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.LaoFix = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
