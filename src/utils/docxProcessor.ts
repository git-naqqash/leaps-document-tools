import JSZip from "jszip";

const wNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// Schema-defined element order to ensure strict compliance with OpenXML specification
const pPrOrder = [
  "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
  "widowControl", "numPr", "pBdr", "shd", "tabs", "spacing", "ind",
  "contextualSpacing", "mirrorMargins", "textboxTightWrap",
  "suppressPaperActive", "wordWrap", "overflowPunct",
  "topLinePunct", "autoSpaceDE", "autoSpaceDN", "adjustRightInd",
  "snapToGrid", "rPr", "sectPr", "pPrChange"
];

const rPrOrder = [
  "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
  "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
  "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
  "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
  "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
  "eastAsianLayout", "specVanish", "oMath"
];

// Helper to find a child by its local name (ignoring namespaces for browser compatibility)
export function findChildByLocalName(element: Element, localName: string): Element | null {
  for (let i = 0; i < element.children.length; i++) {
    if (element.children[i].localName === localName) {
      return element.children[i];
    }
  }
  return null;
}

// Helper to get all descendant elements of a specific local name
export function getDescendantsByLocalName(element: Element, localName: string): Element[] {
  const result: Element[] = [];
  const list = element.getElementsByTagName("*");
  for (let i = 0; i < list.length; i++) {
    if (list[i].localName === localName) {
      result.push(list[i]);
    }
  }
  return result;
}

// Helper to get text content from a paragraph node by combining all w:t nodes
export function getParagraphText(p: Element): string {
  const tNodes = getDescendantsByLocalName(p, "t");
  return tNodes.map(t => t.textContent || "").join("");
}

// Insert new element into its correct schema position inside parent
function insertInOrder(parent: Element, newChild: Element, order: string[]) {
  const newName = newChild.localName || "";
  const newIdx = order.indexOf(newName);
  if (newIdx === -1) {
    parent.appendChild(newChild);
    return;
  }

  for (let i = 0; i < parent.children.length; i++) {
    const child = parent.children[i];
    const childIdx = order.indexOf(child.localName || "");
    if (childIdx !== -1 && childIdx > newIdx) {
      parent.insertBefore(newChild, child);
      return;
    }
  }

  parent.appendChild(newChild);
}

// Check if a paragraph has the Heading1 style (case-insensitive check)
export function isHeading1(p: Element): boolean {
  const pPr = findChildByLocalName(p, "pPr");
  if (!pPr) return false;
  const pStyle = findChildByLocalName(pPr, "pStyle");
  if (!pStyle) return false;
  const val = pStyle.getAttribute("w:val") || pStyle.getAttribute("val") || "";
  const lowerVal = val.toLowerCase();
  return lowerVal === "heading1" || lowerVal === "heading 1";
}

// Check if a paragraph has the Heading2 style (case-insensitive check)
export function isHeading2(p: Element): boolean {
  const pPr = findChildByLocalName(p, "pPr");
  if (!pPr) return false;
  const pStyle = findChildByLocalName(pPr, "pStyle");
  if (!pStyle) return false;
  const val = pStyle.getAttribute("w:val") || pStyle.getAttribute("val") || "";
  const lowerVal = val.toLowerCase();
  return lowerVal === "heading2" || lowerVal === "heading 2";
}

// Checks if a paragraph is styled or written as a list bullet
export function isBulletOrListItem(p: Element): boolean {
  const pPr = findChildByLocalName(p, "pPr");
  if (pPr && findChildByLocalName(pPr, "numPr")) {
    return true;
  }
  const text = getParagraphText(p).trim();
  // Starts with list/bullet indicators: •, -, *, ●, ○
  return /^[•\-*●○]/.test(text);
}

// Checks if a paragraph is explicitly indented in its XML properties (left, start, firstLine, hanging)
export function isIndented(p: Element): boolean {
  const pPr = findChildByLocalName(p, "pPr");
  if (!pPr) return false;
  const ind = findChildByLocalName(pPr, "ind");
  if (ind) {
    const hasLeft = ind.hasAttribute("w:left") || ind.hasAttribute("left");
    const hasStart = ind.hasAttribute("w:start") || ind.hasAttribute("start");
    const hasFirstLine = ind.hasAttribute("w:firstLine") || ind.hasAttribute("firstLine");
    const hasHanging = ind.hasAttribute("w:hanging") || ind.hasAttribute("hanging");

    if (hasLeft || hasStart || hasFirstLine || hasHanging) {
      return true;
    }
  }
  return false;
}

// Convert paragraph to Heading 2 style and set Bold to false in compliance with XML schemas
export function applyH2AndUnbold(doc: Document, p: Element) {
  // 1. Get or create pPr
  let pPr = findChildByLocalName(p, "pPr");
  if (!pPr) {
    pPr = doc.createElementNS(wNamespace, "w:pPr");
    p.insertBefore(pPr, p.firstChild);
  }

  // 2. Get or create pStyle, set value to Heading2, and insert in schema order
  let pStyle = findChildByLocalName(pPr, "pStyle");
  if (!pStyle) {
    pStyle = doc.createElementNS(wNamespace, "w:pStyle");
    insertInOrder(pPr, pStyle, pPrOrder);
  }
  pStyle.setAttributeNS(wNamespace, "w:val", "Heading2");

  // 3. Process all run elements (w:r) to turn bold off
  const runs = getDescendantsByLocalName(p, "r");
  for (const r of runs) {
    let rPr = findChildByLocalName(r, "rPr");
    if (!rPr) {
      rPr = doc.createElementNS(wNamespace, "w:rPr");
      r.insertBefore(rPr, r.firstChild);
    }

    // Set bold (w:b) to false
    let b = findChildByLocalName(rPr, "b");
    if (!b) {
      b = doc.createElementNS(wNamespace, "w:b");
      insertInOrder(rPr, b, rPrOrder);
    }
    b.setAttributeNS(wNamespace, "w:val", "false");

    // Set bold for complex scripts (w:bCs) to false
    let bCs = findChildByLocalName(rPr, "bCs");
    if (!bCs) {
      bCs = doc.createElementNS(wNamespace, "w:bCs");
      insertInOrder(rPr, bCs, rPrOrder);
    }
    bCs.setAttributeNS(wNamespace, "w:val", "false");
  }
}

// Helper to perform conservative normalization for exact heading matching
// Helper to perform conservative normalization for exact heading matching
export function normalizeForMatching(str: string): string {
  if (!str) return "";
  
  // 1. Remove zero-width characters
  let normalized = str.replace(/[\u200B-\u200D\uFEFF]/g, "");

  // 2. Treat tab spacing as equivalent to space
  normalized = normalized.replace(/\t/g, " ");

  // 3. Remove leading bullet characters and any spaces after them
  normalized = normalized.replace(/^\s*[•●○▪▫*+>#\-–—]\s*/, "");

  // 4. Remove trailing colons
  normalized = normalized.replace(/:+$/, "");

  // 5. Treat hyphen, en dash (–), and em dash (—) as equivalent
  normalized = normalized.replace(/[\u2013\u2014-]/g, "-");

  // 6. Treat smart quotes and straight quotes as equivalent
  // Double quotes
  normalized = normalized.replace(/[\u201C\u201D\u201E]/g, '"');
  // Single quotes / apostrophes
  normalized = normalized.replace(/[\u2018\u2019\u201A]/g, "'");

  // 7. Collapse multiple spaces into one
  normalized = normalized.replace(/\s+/g, " ");

  // 8. Recipe serial number equivalence:
  // e.g., "1 - Green Beans", "1 – Green Beans", "1. Green Beans", "1) Green Beans"
  // Normalize prefix to "1 Green Beans" format
  normalized = normalized.replace(/^(\d+)\s*[\.\-\)]\s*/, "$1 ");

  // Case-insensitivity (lowercase) and trim
  return normalized.trim().toLowerCase();
}

// Clean trailing page numbers and dot leaders from outline/TOC lines
export function cleanOutlineLine(line: string): string {
  let cleaned = line.trim();
  // Strip trailing dot leaders, dashes, underscores, spaces, or tabs followed by a page number
  cleaned = cleaned.replace(/[\s\t\.\-_]{2,}\d+$/, "");
  // Strip trailing tab character followed by a page number
  cleaned = cleaned.replace(/\t\d+$/, "");
  return cleaned.trim();
}

// Checks if a line contains markers typical of a TOC page line (dots, page numbers with tabs)
export function isTOCParagraphText(text: string): boolean {
  const trimmed = text.trim();
  // Contains dot leaders (3 or more dots) or underscores (3 or more underscores)
  if (/\.{3,}/.test(trimmed) || /_{3,}/.test(trimmed)) {
    return true;
  }
  // Ends with a page number preceded by a tab, multiple spaces, or dots
  if (/[\s\t\.\-_]\d+$/.test(trimmed)) {
    if (/(\t|\s{2,}|\.{2,})\d+$/.test(trimmed)) {
      return true;
    }
  }
  return false;
}

// Checks if a paragraph element represents a TOC entry based on style or text
export function isTOCParagraph(p: Element): boolean {
  const pPr = findChildByLocalName(p, "pPr");
  if (pPr) {
    const pStyle = findChildByLocalName(pPr, "pStyle");
    if (pStyle) {
      const val = (pStyle.getAttribute("w:val") || pStyle.getAttribute("val") || "").toLowerCase();
      if (val.startsWith("toc") || val.includes("toc")) {
        return true;
      }
    }
  }
  const text = getParagraphText(p);
  return isTOCParagraphText(text);
}

// Main logic to parse a docx archive, identify headings, and rewrite word/document.xml
// Main logic to parse a docx archive, identify headings, and rewrite word/document.xml
export async function processDocxHeadings(
  docxFile: File,
  outlineLines: { text: string; isHeading1?: boolean }[] | null
): Promise<{
  blob: Blob;
  convertedCount: number;
  unmatchedLines: string[];
  skippedH1Count: number;
  warnings: string[];
  totalOutlineTargets: number;
  convertedParagraphs: string[];
  unmatchedOutlineTargets: string[];
  skippedAmbiguousMatches: { text: string; reason: string }[];
  skippedOutlineTargets: string[];
}> {
  const arrayBuffer = await docxFile.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);

  const documentXmlPath = "word/document.xml";
  const docXmlText = await zip.file(documentXmlPath)?.async("text");

  if (!docXmlText) {
    throw new Error("Invalid DOCX structure: Missing 'word/document.xml' file.");
  }

  const parser = new DOMParser();
  const doc = parser.parseFromString(docXmlText, "text/xml");
  const allPs = getDescendantsByLocalName(doc.documentElement, "p");

  let convertedCount = 0;
  let skippedH1Count = 0;
  const warnings: string[] = [];
  const unmatchedLines: string[] = [];

  // Report collections
  const convertedParagraphs: string[] = [];
  const skippedAmbiguousMatches: { text: string; reason: string }[] = [];
  const unmatchedOutlineTargets: string[] = [];
  const skippedOutlineTargets: string[] = [];

  // Parse and clean outline lines
  const cleanOutlineLines: string[] = [];
  
  // Find first non-empty line
  let firstNonEmptyLine: string | null = null;
  if (outlineLines) {
    for (const item of outlineLines) {
      if (item.text.trim().length > 0) {
        firstNonEmptyLine = item.text.trim();
        break;
      }
    }
  }

  if (outlineLines) {
    for (const item of outlineLines) {
      const rawText = item.text;
      const trimmed = rawText.trim();
      if (trimmed.length === 0) continue;

      // Skip the book title if it is the first non-empty line
      const isBookTitle = firstNonEmptyLine && trimmed === firstNonEmptyLine;

      // Skip rules based on requested patterns
      const isContents = /^contents$/i.test(trimmed);
      const isIntro = /^introduction$/i.test(trimmed);
      const isConclusion = /^conclusion$/i.test(trimmed);
      const isBonus = /^bonus$/i.test(trimmed);
      const isPart = /^part\s+\d+/i.test(trimmed);
      const isChapterPattern1 = /^chapter\s+\d+/i.test(trimmed);
      const isChapterPattern2 = /^chapter\s+\d+\s*[–-]/i.test(trimmed);
      const isChapterPattern3 = /^CHAPTER\s+\d+/i.test(trimmed);

      // TOC page-number lines (ending in page numbers/dot leaders or tab followed by page number)
      const isTOC = /^.+\t\d+$/.test(trimmed) || isTOCParagraphText(trimmed);

      // Heading 1 style from outline file
      const isOutlineH1 = !!item.isHeading1;

      if (
        isBookTitle ||
        isContents ||
        isIntro ||
        isConclusion ||
        isBonus ||
        isPart ||
        isChapterPattern1 ||
        isChapterPattern2 ||
        isChapterPattern3 ||
        isTOC ||
        isOutlineH1
      ) {
        skippedOutlineTargets.push(rawText);
        continue;
      }

      // Clean dot leaders and page numbers
      const cleanedText = cleanOutlineLine(rawText);
      if (cleanedText.length < 3) {
        skippedOutlineTargets.push(rawText);
        continue;
      }

      // Skip if it is just a number
      if (/^\d+$/.test(cleanedText)) {
        skippedOutlineTargets.push(rawText);
        continue;
      }

      cleanOutlineLines.push(cleanedText);
    }
  }

  const totalOutlineTargets = cleanOutlineLines.length;

  if (totalOutlineTargets === 0) {
    throw new Error("No valid outline targets found. Please upload the correct outline file.");
  }

  // --- OUTLINE MODE (Outline is the ONLY source of truth) ---
  
  // Map normalized string to original line for reporting
  const normalizedToOriginalOutline = new Map<string, string>();
  const sanitizedOutlineLines: string[] = [];
  
  for (const line of cleanOutlineLines) {
    const norm = normalizeForMatching(line);
    if (norm.length >= 3) {
      if (!normalizedToOriginalOutline.has(norm)) {
        normalizedToOriginalOutline.set(norm, line);
        sanitizedOutlineLines.push(norm);
      }
    }
  }

  const matchedOutlineNormalized = new Set<string>();

  for (let i = 0; i < allPs.length; i++) {
    const p = allPs[i];
    const text = getParagraphText(p);
    const trimmedText = text.trim();

    if (trimmedText.length === 0) continue;

    const normalizedPText = normalizeForMatching(text);
    if (normalizedPText.length === 0) continue;

    // Check if it matches an item in the outline
    if (normalizedToOriginalOutline.has(normalizedPText)) {
      // Strict blocklist check: do not convert common labels unless they are explicitly in the outline
      const strictBlocklist = [
        "intro", "intro:",
        "introduction", "introduction:",
        "ingredients", "ingredients:",
        "instructions", "instructions:",
        "directions", "directions:",
        "prep time", "prep time:",
        "cook time", "cook time:",
        "storage & shelf life", "storage & shelf life:",
        "nutrition", "nutrition:",
        "canner's tip", "canner's tip:",
        "makes", "makes:",
        "preparation & processing time", "preparation & processing time:",
        "meal composition", "meal composition:",
        "recommended recipes", "recommended recipes:",
        "what to watch", "what to watch:",
        "focus areas", "focus areas:"
      ];
      const lowerTrimmed = trimmedText.toLowerCase();
      if (strictBlocklist.includes(lowerTrimmed)) {
        const hasExactInOutline = cleanOutlineLines.some(
          line => line.toLowerCase() === lowerTrimmed
        );
        if (!hasExactInOutline) {
          skippedAmbiguousMatches.push({ text: trimmedText, reason: "Blocklisted label (not explicitly in outline)" });
          continue;
        }
      }

      // Evaluate guards
      const isH1 = isHeading1(p);
      const isTOC = isTOCParagraph(p);

      const wordCount = trimmedText.split(/\s+/).filter(Boolean).length;
      const isLong = trimmedText.length > 120;
      const hasTooManyWords = wordCount > 15;
      const endsWithPeriod = trimmedText.endsWith(".");
      const failedLengthGuard = isLong || hasTooManyWords || endsWithPeriod;

      if (isH1) {
        skippedH1Count++;
        skippedAmbiguousMatches.push({ text: trimmedText, reason: "Already Heading 1" });
      } else if (isTOC) {
        skippedAmbiguousMatches.push({ text: trimmedText, reason: "Table of Contents line" });
      } else if (failedLengthGuard) {
        let reason = "Length & Punctuation Guard (";
        if (isLong) reason += `length: ${trimmedText.length} chars > 120`;
        else if (hasTooManyWords) reason += `word count: ${wordCount} > 15`;
        else if (endsWithPeriod) reason += "ends with period";
        reason += ")";
        skippedAmbiguousMatches.push({ text: trimmedText, reason });
      } else {
        // Convert paragraph to Heading 2
        applyH2AndUnbold(doc, p);
        convertedCount++;
        convertedParagraphs.push(trimmedText);
        matchedOutlineNormalized.add(normalizedPText);
      }
    }
  }

  // Determine unmatched outline lines
  for (const norm of sanitizedOutlineLines) {
    if (!matchedOutlineNormalized.has(norm)) {
      const orig = normalizedToOriginalOutline.get(norm) || norm;
      unmatchedOutlineTargets.push(orig);
      unmatchedLines.push(orig);
    }
  }

  if (convertedCount === 0 && warnings.length === 0) {
    warnings.push("No H2 headings were converted. No changes made.");
  }

  // Serialize DOM back to XML string without modifying spacing/trimming in the source w:t nodes
  const serializer = new XMLSerializer();
  const updatedXmlText = serializer.serializeToString(doc);

  // Write updated XML back to ZIP
  zip.file(documentXmlPath, updatedXmlText);

  // Generate output Blob with compression DEFLATE
  const blob = await zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    compression: "DEFLATE",
    compressionOptions: { level: 6 }
  });

  return {
    blob,
    convertedCount,
    unmatchedLines,
    skippedH1Count,
    warnings,
    totalOutlineTargets,
    convertedParagraphs,
    unmatchedOutlineTargets,
    skippedAmbiguousMatches,
    skippedOutlineTargets
  };
}
