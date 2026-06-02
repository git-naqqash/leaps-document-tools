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

// Helper to check if a paragraph at a given index is followed by recipe content
export function isRecipeHeading(allPs: Element[], index: number): boolean {
  let checkIndex = index + 1;
  let nonMetadataCount = 0;

  // Check next 5 paragraphs
  while (checkIndex < allPs.length && nonMetadataCount < 5) {
    const nextP = allPs[checkIndex];
    const nextText = getParagraphText(nextP).trim();
    if (nextText.length > 0) {
      // If the next paragraph is a Heading 1 or Heading 2, the section has ended
      if (isHeading1(nextP) || isHeading2(nextP)) {
        break;
      }
      nonMetadataCount++;
      
      // Key terms indicating a recipe follows
      const hasKeywords = /^(ingredients|directions|instructions|method|preparation|servings|prep\s*time|cook\s*time|nutrition|calories)/i.test(nextText);
      
      // Ingredient measurements list (e.g. 1 cup, 1/2 tsp, 2 large, etc.)
      const hasIngredientsList = /^\d+[\s\d\/¼½¾.,-]*\s*(cup|tbsp|tsp|gram|g|oz|ml|pound|lb|pinch|can|clove|slice|teaspoon|tablespoon|piece|pkg|large|medium|small|head|cloves|slices|cans|grams|ounces|ml|cups|tablespoons|teaspoons)/i.test(nextText);
      
      // Numbered cooking instructions (e.g. 1. Mix, Step 2: Heat)
      const hasNumberedSteps = /^(step\s*\d+|\d+\s*[\.\)-])/i.test(nextText);

      if (hasKeywords || hasIngredientsList || hasNumberedSteps) {
        return true;
      }
    }
    checkIndex++;
  }
  return false;
}

// Main logic to parse a docx archive, identify headings, and rewrite word/document.xml
export async function processDocxHeadings(
  docxFile: File,
  outlineLines: string[] | null,
  isRecipeBook: boolean
): Promise<{
  blob: Blob;
  convertedCount: number;
  recipeConvertedCount: number;
  subtopicConvertedCount: number;
  unmatchedLines: string[];
  skippedH1Count: number;
  warnings: string[];
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
  let recipeConvertedCount = 0;
  let subtopicConvertedCount = 0;
  let skippedH1Count = 0;
  const unmatchedLines: string[] = [];
  const warnings: string[] = [];

  const cleanOutlineLines = outlineLines
    ? outlineLines.map(line => line.trim()).filter(line => line.length > 0)
    : null;

  if (cleanOutlineLines) {
    // --- OPTION 1: OUTLINE MODE ---
    // Cache text and H1 details of all document paragraphs
    const paragraphsInfo = allPs.map((p, idx) => ({
      index: idx,
      pElement: p,
      text: getParagraphText(p).trim(),
      isH1: isHeading1(p)
    })).filter(info => info.text.length > 0);

    const convertedIndices = new Set<number>();

    for (const line of cleanOutlineLines) {
      let isLineMatched = false;

      // 1. Try exact case-sensitive match
      const caseSensitiveMatches = paragraphsInfo.filter(info => info.text === line);

      if (caseSensitiveMatches.length > 0) {
        isLineMatched = true;
        for (const match of caseSensitiveMatches) {
          if (match.isH1) {
            skippedH1Count++;
          } else {
            if (!convertedIndices.has(match.index)) {
              applyH2AndUnbold(doc, match.pElement);
              convertedCount++;
              if (isRecipeBook) {
                if (isRecipeHeading(allPs, match.index)) {
                  recipeConvertedCount++;
                } else {
                  subtopicConvertedCount++;
                }
              }
              convertedIndices.add(match.index);
            }
          }
        }
      } else {
        // 2. Try safe fallback case-insensitive match only if exactly one matches
        const caseInsensitiveMatches = paragraphsInfo.filter(
          info => info.text.toLowerCase() === line.toLowerCase()
        );

        if (caseInsensitiveMatches.length === 1) {
          isLineMatched = true;
          const match = caseInsensitiveMatches[0];
          if (match.isH1) {
            skippedH1Count++;
          } else {
            if (!convertedIndices.has(match.index)) {
              applyH2AndUnbold(doc, match.pElement);
              convertedCount++;
              if (isRecipeBook) {
                if (isRecipeHeading(allPs, match.index)) {
                  recipeConvertedCount++;
                } else {
                  subtopicConvertedCount++;
                }
              }
              convertedIndices.add(match.index);
            }
          }
        }
      }

      if (!isLineMatched) {
        unmatchedLines.push(line);
      }
    }
  } else {
    // --- OPTION 2: AUTOMATIC DETECTION MODE ---
    for (let i = 0; i < allPs.length; i++) {
      const p = allPs[i];
      const text = getParagraphText(p).trim();

      // Skip empty lines and H1 paragraphs
      if (text.length === 0 || isHeading1(p)) {
        if (text.length > 0 && isHeading1(p)) {
          // If we encounter a Heading 1 under automatic mode, count it for the skipped logs
          skippedH1Count++;
        }
        continue;
      }

      const isBullet = isBulletOrListItem(p);
      const isInd = isIndented(p);
      const isAlreadyH2 = isHeading2(p);

      // Automatic mode must skip bullets, indents, and existing H2s
      if (!isBullet && !isInd && !isAlreadyH2) {
        let shouldConvert = false;
        let isRecipe = false;

        if (isRecipeBook) {
          // Recipe Book Heuristics
          if (isRecipeHeading(allPs, i)) {
            shouldConvert = true;
            isRecipe = true;
          } else if (checkNonRecipeHeuristic(allPs, i)) {
            shouldConvert = true;
            isRecipe = false;
          }
        } else {
          // Non-Recipe Book Heuristics
          shouldConvert = checkNonRecipeHeuristic(allPs, i);
        }

        if (shouldConvert) {
          applyH2AndUnbold(doc, p);
          convertedCount++;
          if (isRecipeBook) {
            if (isRecipe) {
              recipeConvertedCount++;
            } else {
              subtopicConvertedCount++;
            }
          }
        }
      }
    }
  }

  if (convertedCount === 0) {
    warnings.push("No clear H2 candidates found. No changes made.");
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
  return { blob, convertedCount, recipeConvertedCount, subtopicConvertedCount, unmatchedLines, skippedH1Count, warnings };
}

// Conservative check for Non-Recipe headings
function checkNonRecipeHeuristic(allPs: Element[], index: number): boolean {
  const p = allPs[index];
  const text = getParagraphText(p).trim();

  // Basic title length bounds
  if (text.length < 3 || text.length > 80) return false;
  
  // Word count check (1 to 12 words)
  const wordCount = text.split(/\s+/).length;
  if (wordCount > 12) return false;

  // Titles should not end with terminal punctuation
  if (/[.?!:]$/.test(text)) return false;

  // Must not start with a lowercase letter (should be capitalized)
  if (/^[a-z]/.test(text)) return false;

  // Look ahead for body text belonging to this topic
  let nextText = "";
  let nextP: Element | null = null;
  for (let j = index + 1; j < allPs.length; j++) {
    const t = getParagraphText(allPs[j]).trim();
    if (t.length > 0) {
      nextText = t;
      nextP = allPs[j];
      break;
    }
  }

  if (nextText.length > 0 && nextP) {
    // Check if next paragraph represents standard body text
    const nextWordCount = nextText.split(/\s+/).length;
    const nextEndsWithPunct = /[.?!]$/.test(nextText);
    const nextIsHeading = isHeading1(nextP) || isHeading2(nextP);

    // If next paragraph is not a heading and has body characteristics (long or ends in punctuation)
    if (!nextIsHeading && (nextWordCount > 15 || nextEndsWithPunct)) {
      return true;
    }
  }

  return false;
}
