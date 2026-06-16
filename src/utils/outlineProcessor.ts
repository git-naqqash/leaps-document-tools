import JSZip from "jszip";
import { getDescendantsByLocalName, getParagraphText, isHeading1 } from "./docxProcessor";

export async function parseOutlineFile(file: File): Promise<{ text: string; isHeading1: boolean }[]> {
  const extension = file.name.split(".").pop()?.toLowerCase();

  if (extension === "txt") {
    // Read text file line by line
    const text = await file.text();
    return text
      .split(/\r?\n/)
      .map(line => line.trim())
      .filter(line => line.length > 0)
      .map(line => ({ text: line, isHeading1: false }));
  } else if (extension === "docx") {
    // Read DOCX archive using JSZip and DOMParser
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);
    
    const documentXmlPath = "word/document.xml";
    const docXmlText = await zip.file(documentXmlPath)?.async("text");
    
    if (!docXmlText) {
      throw new Error("Invalid outline DOCX structure: Missing 'word/document.xml'.");
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(docXmlText, "text/xml");
    const paragraphs = getDescendantsByLocalName(doc.documentElement, "p");

    return paragraphs
      .map(p => ({
        text: getParagraphText(p).trim(),
        isHeading1: isHeading1(p)
      }))
      .filter(item => item.text.length > 0);
  } else {
    throw new Error("Unsupported outline file format. Please upload a .txt or .docx file.");
  }
}
