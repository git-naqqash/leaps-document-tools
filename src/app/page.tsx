"use client";

import React, { useState, useRef } from "react";
import { processDocxHeadings } from "../utils/docxProcessor";
import { parseOutlineFile } from "../utils/outlineProcessor";
import { extractImagesFromFiles } from "../utils/imageExtractor";

// Custom Inline SVG Icons for premium styling
const FileWordIcon = () => (
  <svg className="w-8 h-8 text-blue-600" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
  </svg>
);

const FileOutlineIcon = () => (
  <svg className="w-6 h-6 text-slate-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M4 6h16M4 12h16M4 18h7" />
  </svg>
);

const FileZipIcon = () => (
  <svg className="w-8 h-8 text-amber-600" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M8 4H6a2 2 0 00-2 2v12a2 2 0 002 2h12a2 2 0 002-2V6a2 2 0 00-2-2h-2m-4-1v8m0 0l3-3m-3 3L9 8m-5 5h2.586a1 1 0 01.707.293l2.414 2.414a1 1 0 00.707.293h3.172a1 1 0 00.707-.293l2.414-2.414a1 1 0 01.707-.293H20" />
  </svg>
);

const DownloadIcon = () => (
  <svg className="w-5 h-5 mr-2" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
  </svg>
);

const CheckCircleIcon = () => (
  <svg className="w-5 h-5 text-emerald-500 mr-2 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
  </svg>
);

const ExclamationIcon = () => (
  <svg className="w-5 h-5 text-amber-500 mr-2 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
  </svg>
);

const ErrorIcon = () => (
  <svg className="w-5 h-5 text-rose-500 mr-2 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M10 14l2-2m0 0l2-2m-2 2l-2-2m2 2l2 2m7-2a9 9 0 11-18 0 9 9 0 0118 0z" />
  </svg>
);

const TrashIcon = () => (
  <svg className="w-4 h-4 text-slate-400 hover:text-rose-500 transition-colors" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
  </svg>
);

function isSameFile(mainFile: File | null, outlineFile: File | null): boolean {
  if (!mainFile || !outlineFile) return false;

  return (
    mainFile.name === outlineFile.name ||
    (
      mainFile.name === outlineFile.name &&
      mainFile.size === outlineFile.size
    ) ||
    (
      mainFile.name === outlineFile.name &&
      mainFile.size === outlineFile.size &&
      mainFile.lastModified === outlineFile.lastModified
    )
  );
}

const formatSize = (bytes: number) => {
  if (bytes === 0) return "0 B";
  const k = 1024;
  const sizes = ["B", "KB", "MB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + " " + sizes[i];
};

export default function Home() {
  // --- STATE FOR NON-RECIPE BOOK H2 MAKER (1A) ---
  const [nrOriginalSize, setNrOriginalSize] = useState(0);
  const [nrOutputSize, setNrOutputSize] = useState(0);
  const [nrMainFile, setNrMainFile] = useState<File | null>(null);
  const [nrOutlineFile, setNrOutlineFile] = useState<File | null>(null);
  const [nrOutlineText, setNrOutlineText] = useState("");
  const [nrStatus, setNrStatus] = useState<"idle" | "processing" | "success" | "error">("idle");
  const [nrMessage, setNrMessage] = useState("");
  const [nrConvertedCount, setNrConvertedCount] = useState(0);
  const [nrWarnings, setNrWarnings] = useState<string[]>([]);
  const [nrUnmatchedLines, setNrUnmatchedLines] = useState<string[]>([]);
  const [nrSkippedH1Count, setNrSkippedH1Count] = useState(0);
  const [nrOutlineLinesCount, setNrOutlineLinesCount] = useState(0);
  const [nrUsedOutline, setNrUsedOutline] = useState(false);

  // Detailed report states for Non-Recipe Book
  const [nrConvertedParagraphs, setNrConvertedParagraphs] = useState<string[]>([]);
  const [nrUnmatchedOutlineTargets, setNrUnmatchedOutlineTargets] = useState<string[]>([]);
  const [nrSkippedAmbiguousMatches, setNrSkippedAmbiguousMatches] = useState<{ text: string; reason: string }[]>([]);
  const [nrTotalOutlineTargets, setNrTotalOutlineTargets] = useState(0);
  const [nrOutlineSource, setNrOutlineSource] = useState("");
  const [nrSkippedOutlineTargets, setNrSkippedOutlineTargets] = useState<string[]>([]);

  const nrMainInputRef = useRef<HTMLInputElement>(null);
  const nrOutlineInputRef = useRef<HTMLInputElement>(null);

  // --- STATE FOR RECIPE BOOK H2 MAKER (1B) ---
  const [rOriginalSize, setROriginalSize] = useState(0);
  const [rOutputSize, setROutputSize] = useState(0);
  const [rMainFile, setRMainFile] = useState<File | null>(null);
  const [rOutlineFile, setROutlineFile] = useState<File | null>(null);
  const [rOutlineText, setROutlineText] = useState("");
  const [rStatus, setRStatus] = useState<"idle" | "processing" | "success" | "error">("idle");
  const [rMessage, setRMessage] = useState("");
  const [rConvertedCount, setRConvertedCount] = useState(0);
  const [rWarnings, setRWarnings] = useState<string[]>([]);
  const [rUnmatchedLines, setRUnmatchedLines] = useState<string[]>([]);
  const [rSkippedH1Count, setRSkippedH1Count] = useState(0);
  const [rOutlineLinesCount, setROutlineLinesCount] = useState(0);
  const [rUsedOutline, setRUsedOutline] = useState(false);

  // Detailed report states for Recipe Book
  const [rConvertedParagraphs, setRConvertedParagraphs] = useState<string[]>([]);
  const [rUnmatchedOutlineTargets, setRUnmatchedOutlineTargets] = useState<string[]>([]);
  const [rSkippedAmbiguousMatches, setRSkippedAmbiguousMatches] = useState<{ text: string; reason: string }[]>([]);
  const [rTotalOutlineTargets, setRTotalOutlineTargets] = useState(0);
  const [rOutlineSource, setROutlineSource] = useState("");
  const [rSkippedOutlineTargets, setRSkippedOutlineTargets] = useState<string[]>([]);

  const rMainInputRef = useRef<HTMLInputElement>(null);
  const rOutlineInputRef = useRef<HTMLInputElement>(null);

  // --- STATE FOR IMAGE EXTRACTOR (2) ---
  const [ieFiles, setIeFiles] = useState<File[]>([]);
  const [ieStatus, setIeStatus] = useState<"idle" | "processing" | "success" | "error">("idle");
  const [ieMessage, setIeMessage] = useState("");
  const [ieLogs, setIeLogs] = useState<string[]>([]);
  const [ieWarnings, setIeWarnings] = useState<string[]>([]);
  const ieInputRef = useRef<HTMLInputElement>(null);

  // Helper to trigger browser download
  const triggerDownload = (fileOrBlob: Blob | File, filename: string) => {
    const url = URL.createObjectURL(fileOrBlob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  // --- HANDLERS FOR NON-RECIPE BOOK H2 MAKER (1A) ---
  const handleNrMainChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] || null;
    if (file && !file.name.toLowerCase().endsWith(".docx")) {
      setNrStatus("error");
      setNrMessage("Unsupported file type. Non-Recipe H2 Maker only accepts .docx files.");
      setNrMainFile(null);
      return;
    }
    setNrMainFile(file);
    if (file) {
      setNrStatus("idle");
      setNrMessage("");
    }
  };

  const handleNrOutlineChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] || null;
    if (file && !file.name.toLowerCase().endsWith(".docx") && !file.name.toLowerCase().endsWith(".txt")) {
      setNrStatus("error");
      setNrMessage("Unsupported outline file type. Optional Outline File accepts .docx and .txt only.");
      setNrOutlineFile(null);
      return;
    }
    setNrOutlineFile(file);
    if (file) {
      setNrStatus("idle");
      setNrMessage("");
    }
  };

  const runNonRecipeH2 = async () => {
    if (!nrMainFile) {
      setNrStatus("error");
      setNrMessage("Please upload the main DOCX file.");
      return;
    }

    const manualOutlineText = nrOutlineText.trim();
    if (!nrOutlineFile && !manualOutlineText) {
      setNrStatus("error");
      setNrMessage("Please upload or paste a separate outline first. Outline is required.");
      return;
    }

    if (isSameFile(nrMainFile, nrOutlineFile)) {
      setNrStatus("error");
      setNrMessage("Error: The outline file cannot be the same as the main document. Please upload a separate outline file.");
      return;
    }

    setNrStatus("processing");
    setNrMessage("Reading document and processing Heading 2 nodes...");
    setNrWarnings([]);
    setNrOriginalSize(nrMainFile.size);
    setNrUnmatchedLines([]);
    setNrSkippedH1Count(0);
    setNrOutlineLinesCount(0);
    setNrUsedOutline(false);

    // Clear previous report stats
    setNrConvertedParagraphs([]);
    setNrUnmatchedOutlineTargets([]);
    setNrSkippedAmbiguousMatches([]);
    setNrTotalOutlineTargets(0);
    setNrOutlineSource("");
    setNrSkippedOutlineTargets([]);

    try {
      let outlineLines: { text: string; isHeading1?: boolean }[] | null = null;
      let usedOutline = false;

      // 1. Priority check: manual outline text box first
      const cleanManualLines = nrOutlineText
        .split(/\r?\n/)
        .map(line => line.trim())
        .filter(line => line.length > 0);

      if (cleanManualLines.length > 0) {
        outlineLines = cleanManualLines.map(line => ({ text: line, isHeading1: false }));
        usedOutline = true;
        setNrOutlineLinesCount(cleanManualLines.length);
        setNrUsedOutline(true);
        setNrOutlineSource("Pasted Outline Text");
      } else if (nrOutlineFile) {
        // 2. Outline file upload second
        setNrMessage("Parsing outline file...");
        const fileLines = await parseOutlineFile(nrOutlineFile);
        outlineLines = fileLines;
        usedOutline = true;
        setNrOutlineLinesCount(fileLines.length);
        setNrUsedOutline(true);
        setNrOutlineSource(nrOutlineFile.name);
        if (fileLines.length === 0) {
          setNrWarnings(["Uploaded outline file is empty or has no usable lines."]);
        }
      }

      setNrMessage("Identifying non-recipe heading structures and unbolding...");
      const {
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
      } = await processDocxHeadings(
        nrMainFile,
        outlineLines
      );

      setNrOutputSize(blob.size);
      setNrConvertedCount(convertedCount);
      setNrUnmatchedLines(unmatchedLines);
      setNrSkippedH1Count(skippedH1Count);
      setNrWarnings((prev) => [...prev, ...warnings]);

      // Set the detailed report states
      setNrTotalOutlineTargets(totalOutlineTargets);
      setNrConvertedParagraphs(convertedParagraphs);
      setNrUnmatchedOutlineTargets(unmatchedOutlineTargets);
      setNrSkippedAmbiguousMatches(skippedAmbiguousMatches);
      setNrSkippedOutlineTargets(skippedOutlineTargets);

      if (convertedCount === 0) {
        setNrStatus("success");
        setNrMessage("No headings were converted. No changes made.");
      } else {
        const originalName = nrMainFile.name.replace(/\.docx$/i, "");
        const outputFilename = `${originalName}-non-recipe-H2-updated.docx`;
        triggerDownload(blob, outputFilename);

        setNrStatus("success");
        setNrMessage(`Successfully formatted ${convertedCount} headings. Your updated file has been downloaded.`);
      }
    } catch (error: unknown) {
      setNrStatus("error");
      setNrMessage(error instanceof Error ? error.message : "An unexpected error occurred during processing.");
    }
  };

  // --- HANDLERS FOR RECIPE BOOK H2 MAKER (1B) ---
  const handleRMainChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] || null;
    if (file && !file.name.toLowerCase().endsWith(".docx")) {
      setRStatus("error");
      setRMessage("Unsupported file type. Recipe Book H2 Maker only accepts .docx files.");
      setRMainFile(null);
      return;
    }
    setRMainFile(file);
    if (file) {
      setRStatus("idle");
      setRMessage("");
    }
  };

  const handleROutlineChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] || null;
    if (file && !file.name.toLowerCase().endsWith(".docx") && !file.name.toLowerCase().endsWith(".txt")) {
      setRStatus("error");
      setRMessage("Unsupported outline file type. Optional Outline File accepts .docx and .txt only.");
      setROutlineFile(null);
      return;
    }
    setROutlineFile(file);
    if (file) {
      setRStatus("idle");
      setRMessage("");
    }
  };

  const runRecipeH2 = async () => {
    if (!rMainFile) {
      setRStatus("error");
      setRMessage("Please upload the main DOCX file.");
      return;
    }

    const manualOutlineText = rOutlineText.trim();
    if (!rOutlineFile && !manualOutlineText) {
      setRStatus("error");
      setRMessage("Please upload or paste a separate outline first. Outline is required.");
      return;
    }

    if (isSameFile(rMainFile, rOutlineFile)) {
      setRStatus("error");
      setRMessage("Error: The outline file cannot be the same as the main document. Please upload a separate outline file.");
      return;
    }

    setRStatus("processing");
    setRMessage("Reading document and analyzing recipe contents...");
    setRWarnings([]);
    setROriginalSize(rMainFile.size);
    setRUnmatchedLines([]);
    setRSkippedH1Count(0);
    setROutlineLinesCount(0);
    setRUsedOutline(false);

    // Clear previous report stats
    setRConvertedParagraphs([]);
    setRUnmatchedOutlineTargets([]);
    setRSkippedAmbiguousMatches([]);
    setRTotalOutlineTargets(0);
    setROutlineSource("");
    setRSkippedOutlineTargets([]);

    try {
      let outlineLines: { text: string; isHeading1?: boolean }[] | null = null;
      let usedOutline = false;

      // 1. Priority check: manual outline text box first
      const cleanManualLines = rOutlineText
        .split(/\r?\n/)
        .map(line => line.trim())
        .filter(line => line.length > 0);

      if (cleanManualLines.length > 0) {
        outlineLines = cleanManualLines.map(line => ({ text: line, isHeading1: false }));
        usedOutline = true;
        setROutlineLinesCount(cleanManualLines.length);
        setRUsedOutline(true);
        setROutlineSource("Pasted Outline Text");
      } else if (rOutlineFile) {
        // 2. Outline file upload second
        setRMessage("Parsing recipe outline file...");
        const fileLines = await parseOutlineFile(rOutlineFile);
        outlineLines = fileLines;
        usedOutline = true;
        setROutlineLinesCount(fileLines.length);
        setRUsedOutline(true);
        setROutlineSource(rOutlineFile.name);
        if (fileLines.length === 0) {
          setRWarnings(["Uploaded outline file is empty or has no usable lines."]);
        }
      }

      setRMessage("Detecting recipe titles, instructions, and list elements...");
      const {
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
      } = await processDocxHeadings(
        rMainFile,
        outlineLines
      );

      setROutputSize(blob.size);
      setRConvertedCount(convertedCount);
      setRUnmatchedLines(unmatchedLines);
      setRSkippedH1Count(skippedH1Count);
      setRWarnings((prev) => [...prev, ...warnings]);

      // Set the detailed report states
      setRTotalOutlineTargets(totalOutlineTargets);
      setRConvertedParagraphs(convertedParagraphs);
      setRUnmatchedOutlineTargets(unmatchedOutlineTargets);
      setRSkippedAmbiguousMatches(skippedAmbiguousMatches);
      setRSkippedOutlineTargets(skippedOutlineTargets);

      if (convertedCount === 0) {
        setRStatus("success");
        setRMessage("No headings were converted. No changes made.");
      } else {
        const originalName = rMainFile.name.replace(/\.docx$/i, "");
        const outputFilename = `${originalName}-recipe-H2-updated.docx`;
        triggerDownload(blob, outputFilename);

        setRStatus("success");
        setRMessage(`Successfully formatted ${convertedCount} headings. Your updated file has been downloaded.`);
      }
    } catch (error: unknown) {
      setRStatus("error");
      setRMessage(error instanceof Error ? error.message : "An unexpected error occurred during processing.");
    }
  };

  // --- HANDLERS FOR IMAGE EXTRACTOR (2) ---
  const handleIeFilesChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFiles = Array.from(e.target.files || []);
    const validFiles = selectedFiles.filter((file) => {
      const ext = file.name.split(".").pop()?.toLowerCase();
      return ext && ["docx", "pptx", "xlsx", "zip"].includes(ext);
    });

    if (validFiles.length !== selectedFiles.length) {
      setIeStatus("error");
      setIeMessage("Some uploaded files were skipped. Only .docx, .pptx, .xlsx, and .zip archives are supported.");
    } else {
      setIeStatus("idle");
      setIeMessage("");
    }

    setIeFiles((prev) => {
      const all = [...prev, ...validFiles];
      return all.filter((v, i, a) => a.findIndex((t) => t.name === v.name && t.size === v.size) === i);
    });
  };

  const removeIeFile = (index: number) => {
    setIeFiles((prev) => prev.filter((_, i) => i !== index));
    if (ieFiles.length <= 1) {
      setIeStatus("idle");
      setIeMessage("");
    }
  };

  const runImageExtraction = async () => {
    if (ieFiles.length === 0) {
      setIeStatus("error");
      setIeMessage("Please upload at least one document (.docx, .pptx, .xlsx) or ZIP file.");
      return;
    }

    setIeStatus("processing");
    setIeLogs([]);
    setIeWarnings([]);
    setIeMessage("Starting media archive analysis...");

    try {
      const { blob, warnings } = await extractImagesFromFiles(ieFiles, (logText) => {
        setIeLogs((prev) => [...prev, logText]);
      });

      setIeWarnings(warnings);
      triggerDownload(blob, "extracted-images-master.zip");
      setIeStatus("success");
      setIeMessage("Image extraction completed! Master ZIP file has been downloaded.");
    } catch (error: unknown) {
      setIeStatus("error");
      setIeMessage(error instanceof Error ? error.message : "An error occurred during image extraction.");
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 font-sans pb-16">
      {/* Decorative Grid Header */}
      <header className="relative overflow-hidden bg-gradient-to-r from-blue-900 to-indigo-950 text-white py-12 px-6 shadow-md">
        <div className="absolute inset-0 bg-[radial-gradient(#3b82f6_1px,transparent_1px)] [background-size:16px_16px] opacity-10"></div>
        <div className="max-w-6xl mx-auto relative z-10 flex flex-col md:flex-row items-center justify-between gap-6">
          <div>
            <h1 className="text-4xl font-extrabold tracking-tight sm:text-5xl">
              LEAPS Document Tools
            </h1>
            <p className="mt-3 text-lg text-blue-200 font-medium">
              Format Word headings and extract images safely in your browser.
            </p>
          </div>
          <div className="bg-white/10 backdrop-blur-md rounded-xl p-4 border border-white/10 text-sm max-w-xs text-blue-100">
            <span className="font-semibold text-white block mb-1">🔒 100% Client-Side Privacy</span>
            Your files never leave your computer. All rendering and extraction processes run purely in browser memory.
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-6 mt-12 space-y-12">
        
        {/* ====================================================
            SECTION 1: H2 MAKER
            ==================================================== */}
        <section className="space-y-6">
          <div className="border-b border-slate-200 pb-3 flex items-center justify-between">
            <h2 className="text-2xl font-bold text-slate-900 flex items-center">
              <span className="w-2.5 h-6 bg-blue-600 rounded-full mr-3 block"></span>
              H2 Maker
            </h2>
            <span className="text-xs text-slate-500 bg-slate-200/60 px-2 py-0.5 rounded font-semibold uppercase tracking-wider">
              DOCX Styles
            </span>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
            
            {/* CARD 1A: NON-RECIPE BOOK H2 MAKER */}
            <div className="bg-white rounded-2xl border border-slate-200/80 p-6 shadow-sm hover:shadow-md transition-all flex flex-col justify-between">
              <div>
                <div className="flex justify-between items-start">
                  <h3 className="text-lg font-bold text-slate-900">Non-Recipe Book H2 Maker</h3>
                  <span className="bg-blue-50 text-blue-700 text-xs px-2.5 py-1 rounded-full font-semibold">General Book</span>
                </div>
                <p className="mt-2 text-sm text-slate-600">
                  For regular nonfiction or fiction books. Converts only clear section/subtopic titles to Heading 2.
                </p>

                {/* Upload Fields Container */}
                <div className="mt-6 space-y-4">
                  {/* Main File */}
                  <div>
                    <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">
                      Main Word File (.docx only)
                    </label>
                    {nrMainFile ? (
                      <div className="flex items-center justify-between bg-blue-50/70 border border-blue-200 rounded-xl p-3">
                        <div className="flex items-center overflow-hidden mr-2">
                          <FileWordIcon />
                          <span className="ml-3 text-sm font-semibold truncate text-blue-900">{nrMainFile.name}</span>
                        </div>
                        <button
                          onClick={() => { setNrMainFile(null); setNrStatus("idle"); }}
                          className="p-1 hover:bg-blue-100 rounded-lg transition"
                          title="Remove file"
                        >
                          <TrashIcon />
                        </button>
                      </div>
                    ) : (
                      <div
                        onClick={() => nrMainInputRef.current?.click()}
                        className="border-2 border-dashed border-slate-200 hover:border-blue-400 hover:bg-slate-50/50 rounded-xl p-6 text-center cursor-pointer transition"
                      >
                        <FileWordIcon />
                        <p className="mt-2 text-sm font-medium text-slate-600">Click to upload main DOCX</p>
                        <p className="text-xs text-slate-400 mt-1">Accepts only .docx</p>
                        <input
                          ref={nrMainInputRef}
                          type="file"
                          accept=".docx"
                          onChange={handleNrMainChange}
                          className="hidden"
                        />
                      </div>
                    )}
                  </div>

                  {/* Manual Outline Text Area */}
                  <div>
                    <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">
                      Paste Outline / Subtopics Here
                    </label>
                    <textarea
                      value={nrOutlineText}
                      onChange={(e) => setNrOutlineText(e.target.value)}
                      placeholder="Paste each H2 heading, recipe name, or subtopic on a separate line."
                      rows={4}
                      className="w-full p-3 text-sm border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent font-mono bg-slate-50/30"
                    />
                  </div>

                  {/* Outline File */}
                  <div>
                    <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">
                      Required Outline File (.DOCX, .TXT)
                    </label>
                    {nrOutlineFile ? (
                      <div className="flex items-center justify-between bg-slate-50 border border-slate-200 rounded-xl p-3">
                        <div className="flex items-center overflow-hidden mr-2">
                          <FileOutlineIcon />
                          <span className="ml-3 text-sm font-medium truncate text-slate-700">{nrOutlineFile.name}</span>
                        </div>
                        <button
                          onClick={() => { setNrOutlineFile(null); setNrStatus("idle"); }}
                          className="p-1 hover:bg-slate-200 rounded-lg transition"
                          title="Remove outline"
                        >
                          <TrashIcon />
                        </button>
                      </div>
                    ) : (
                      <div
                        onClick={() => nrOutlineInputRef.current?.click()}
                        className="border-2 border-dashed border-slate-200 hover:border-slate-400 hover:bg-slate-50/50 rounded-xl p-4 text-center cursor-pointer transition"
                      >
                        <p className="text-xs font-semibold text-slate-600">+ Upload Outline File</p>
                        <p className="text-[10px] text-slate-400 mt-0.5">Supports .docx and .txt</p>
                        <input
                          ref={nrOutlineInputRef}
                          type="file"
                          accept=".docx,.txt"
                          onChange={handleNrOutlineChange}
                          className="hidden"
                        />
                      </div>
                    )}
                    <p className="text-[11px] text-amber-600 mt-2 font-medium">
                      ⚠️ Upload a separate outline file only. Do not upload the same manuscript file here.
                    </p>
                  </div>
                </div>

                {/* Info disclaimer */}
                <div className="mt-4 p-3 bg-slate-50 rounded-xl border border-slate-100 text-xs text-slate-500">
                  Only exact lines from the uploaded or pasted outline will be converted. No automatic heading detection is used.
                </div>

                {/* Status Screen */}
                {nrStatus !== "idle" && (
                  <div className={`mt-6 p-4 rounded-xl text-sm ${
                    nrStatus === "processing" ? "bg-blue-50 border border-blue-100 text-blue-800" :
                    nrStatus === "success" ? "bg-emerald-50 border border-emerald-100 text-emerald-800" :
                    "bg-rose-50 border border-rose-100 text-rose-800"
                  }`}>
                    <div className="flex items-start">
                      {nrStatus === "processing" && (
                        <svg className="animate-spin h-5 w-5 text-blue-600 mr-2 flex-shrink-0" fill="none" viewBox="0 0 24 24">
                          <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                          <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
                        </svg>
                      )}
                      {nrStatus === "success" && <CheckCircleIcon />}
                      {nrStatus === "error" && <ErrorIcon />}
                      <div>
                        <p className="font-semibold">{nrStatus === "processing" ? "Processing..." : nrStatus === "success" ? "Completed" : "Error"}</p>
                        <p className="mt-0.5 text-xs opacity-90">{nrMessage}</p>
                      </div>
                    </div>

                    {nrStatus === "success" && (
                      <div className="mt-3 pt-3 border-t border-slate-200/50 text-xs space-y-3 text-slate-600">
                        <div className="grid grid-cols-2 gap-3 bg-slate-100/60 p-3 rounded-lg border border-slate-200/50 font-medium text-slate-700">
                          <div className="col-span-2"><span className="text-slate-500">Main DOCX file name:</span> <span className="font-bold text-slate-900">{nrMainFile?.name}</span></div>
                          <div className="col-span-2"><span className="text-slate-500">Outline source file name or pasted outline:</span> <span className="font-bold text-slate-900">{nrOutlineSource}</span></div>
                          <div><span className="text-slate-500">Total outline targets extracted:</span> <span className="font-bold text-slate-900">{nrTotalOutlineTargets}</span></div>
                          <div><span className="text-slate-500">Total paragraphs converted by exact outline match:</span> <span className="font-bold text-slate-900">{nrConvertedCount}</span></div>
                          <div><span className="text-slate-500">Total automatic detections used:</span> <span className="font-bold text-slate-900">0</span></div>
                          <div><span className="text-slate-500">Skipped H1 matches:</span> <span className="font-bold text-slate-900">{nrSkippedH1Count}</span></div>
                          <div><span className="text-slate-500">Original file size:</span> <span className="font-bold text-slate-900">{formatSize(nrOriginalSize)}</span></div>
                          <div><span className="text-slate-500">Output file size:</span> <span className="font-bold text-slate-900">{formatSize(nrOutputSize)}</span></div>
                        </div>

                        {nrOutputSize > 2 * nrOriginalSize && (
                          <div className="bg-rose-50 border border-rose-200 text-rose-800 rounded-lg p-2.5 font-semibold text-xs mt-2 flex items-center">
                            <ErrorIcon /> Warning: Output file size increased unexpectedly. Please check compression settings.
                          </div>
                        )}

                        {/* Pre-save / processing report */}
                        <div className="border border-slate-200 rounded-xl p-3 bg-white space-y-3">
                          <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Processing Report</h4>
                          
                          <div className="grid grid-cols-2 gap-2 text-[11px]">
                            <div>Total outline targets found: <span className="font-bold text-slate-800">{nrTotalOutlineTargets}</span></div>
                            <div>Total paragraphs converted: <span className="font-bold text-slate-800">{nrConvertedCount}</span></div>
                          </div>

                          {/* List of converted paragraph texts */}
                          {nrConvertedParagraphs.length > 0 && (
                            <div className="pt-2 border-t border-slate-100">
                              <span className="font-semibold text-slate-700 block mb-1">Converted paragraph list ({nrConvertedParagraphs.length}):</span>
                              <div className="max-h-[100px] overflow-y-auto font-mono text-[10px] text-slate-600 space-y-1 bg-slate-50 p-2 rounded-lg border border-slate-100">
                                {nrConvertedParagraphs.map((text, idx) => (
                                  <div key={idx}>• {text}</div>
                                ))}
                              </div>
                            </div>
                          )}

                          {/* List of outline targets not found */}
                          {nrUnmatchedOutlineTargets.length > 0 && (
                            <div className="pt-2 border-t border-slate-100">
                              <span className="font-semibold text-amber-800 block mb-1">Outline targets not found ({nrUnmatchedOutlineTargets.length}):</span>
                              <div className="max-h-[100px] overflow-y-auto font-mono text-[10px] text-amber-700 space-y-1 bg-amber-50/50 p-2 rounded-lg border border-amber-100">
                                {nrUnmatchedOutlineTargets.map((text, idx) => (
                                  <div key={idx}>• {text}</div>
                                ))}
                              </div>
                            </div>
                          )}

                          {/* List of outline lines skipped because they are H1/title/frontmatter */}
                          {nrSkippedOutlineTargets.length > 0 && (
                            <div className="pt-2 border-t border-slate-100">
                              <span className="font-semibold text-slate-500 block mb-1">Outline lines skipped because they are H1/title/frontmatter ({nrSkippedOutlineTargets.length}):</span>
                              <div className="max-h-[100px] overflow-y-auto font-mono text-[10px] text-slate-500 space-y-1 bg-slate-100 p-2 rounded-lg border border-slate-200">
                                {nrSkippedOutlineTargets.map((text, idx) => (
                                  <div key={idx}>• {text}</div>
                                ))}
                              </div>
                            </div>
                          )}

                          {/* List of skipped ambiguous matches */}
                          {nrSkippedAmbiguousMatches.length > 0 && (
                            <div className="pt-2 border-t border-slate-100">
                              <span className="font-semibold text-rose-800 block mb-1">Skipped Ambiguous Matches ({nrSkippedAmbiguousMatches.length}):</span>
                              <div className="max-h-[100px] overflow-y-auto font-mono text-[10px] text-rose-700 space-y-1 bg-rose-50/30 p-2 rounded-lg border border-rose-100">
                                {nrSkippedAmbiguousMatches.map((item, idx) => (
                                  <div key={idx}>
                                    • <span className="font-semibold">{item.text}</span> <span className="italic text-[9px] text-rose-500">({item.reason})</span>
                                  </div>
                                ))}
                              </div>
                            </div>
                          )}
                        </div>

                        <div className="text-emerald-700 font-semibold flex items-center pt-1 border-t border-slate-200/40">
                          ✓ No text was rewritten. Only Heading 2 style and bold-off formatting were applied.
                        </div>

                        {nrConvertedCount === 0 && nrMainFile && (
                          <button
                            onClick={() => triggerDownload(nrMainFile, nrMainFile.name)}
                            className="mt-3 w-full bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2 px-3 rounded-lg border border-slate-200 transition text-center"
                          >
                            Download unchanged file
                          </button>
                        )}
                      </div>
                    )}

                    {nrWarnings.length > 0 && (
                      <div className="mt-2 text-xs text-amber-700 space-y-1 bg-amber-50 p-2 rounded-lg border border-amber-100">
                        <div className="font-semibold flex items-center"><ExclamationIcon /> Warnings:</div>
                        {nrWarnings.map((w, idx) => (
                          <div key={idx} className="pl-6">• {w}</div>
                        ))}
                      </div>
                    )}
                  </div>
                )}
              </div>

              <div className="mt-8">
                <button
                  onClick={runNonRecipeH2}
                  disabled={
                    nrStatus === "processing" ||
                    !(
                      nrMainFile &&
                      (nrOutlineFile || nrOutlineText.trim().length > 0) &&
                      !isSameFile(nrMainFile, nrOutlineFile)
                    )
                  }
                  className="w-full bg-blue-600 hover:bg-blue-700 disabled:bg-slate-200 disabled:text-slate-400 text-white font-bold py-3 px-4 rounded-xl transition shadow-sm hover:shadow active:scale-[0.99] flex items-center justify-center"
                >
                  Format Non-Recipe Book H2
                </button>
              </div>
            </div>

            {/* CARD 1B: RECIPE BOOK H2 MAKER */}
            <div className="bg-white rounded-2xl border border-slate-200/80 p-6 shadow-sm hover:shadow-md transition-all flex flex-col justify-between">
              <div>
                <div className="flex justify-between items-start">
                  <h3 className="text-lg font-bold text-slate-900">Recipe Book H2 Maker</h3>
                  <span className="bg-indigo-50 text-indigo-700 text-xs px-2.5 py-1 rounded-full font-semibold">Cookbooks</span>
                </div>
                <p className="mt-2 text-sm text-slate-600">
                  For cookbooks and recipe books. Converts recipe names and chapter subtopics to Heading 2 without removing serial numbers.
                </p>

                {/* Upload Fields Container */}
                <div className="mt-6 space-y-4">
                  {/* Main File */}
                  <div>
                    <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">
                      Main Word File (.docx only)
                    </label>
                    {rMainFile ? (
                      <div className="flex items-center justify-between bg-indigo-50/70 border border-indigo-200 rounded-xl p-3">
                        <div className="flex items-center overflow-hidden mr-2">
                          <FileWordIcon />
                          <span className="ml-3 text-sm font-semibold truncate text-indigo-900">{rMainFile.name}</span>
                        </div>
                        <button
                          onClick={() => { setRMainFile(null); setRStatus("idle"); }}
                          className="p-1 hover:bg-indigo-100 rounded-lg transition"
                          title="Remove file"
                        >
                          <TrashIcon />
                        </button>
                      </div>
                    ) : (
                      <div
                        onClick={() => rMainInputRef.current?.click()}
                        className="border-2 border-dashed border-slate-200 hover:border-indigo-400 hover:bg-slate-50/50 rounded-xl p-6 text-center cursor-pointer transition"
                      >
                        <FileWordIcon />
                        <p className="mt-2 text-sm font-medium text-slate-600">Click to upload cookbook DOCX</p>
                        <p className="text-xs text-slate-400 mt-1">Accepts only .docx</p>
                        <input
                          ref={rMainInputRef}
                          type="file"
                          accept=".docx"
                          onChange={handleRMainChange}
                          className="hidden"
                        />
                      </div>
                    )}
                  </div>

                  {/* Manual Outline Text Area */}
                  <div>
                    <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">
                      Paste Outline / Subtopics Here
                    </label>
                    <textarea
                      value={rOutlineText}
                      onChange={(e) => setROutlineText(e.target.value)}
                      placeholder="Paste each H2 heading, recipe name, or subtopic on a separate line."
                      rows={4}
                      className="w-full p-3 text-sm border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-transparent font-mono bg-slate-50/30"
                    />
                  </div>

                  {/* Outline File */}
                  <div>
                    <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">
                      Required Outline File (.DOCX, .TXT)
                    </label>
                    {rOutlineFile ? (
                      <div className="flex items-center justify-between bg-slate-50 border border-slate-200 rounded-xl p-3">
                        <div className="flex items-center overflow-hidden mr-2">
                          <FileOutlineIcon />
                          <span className="ml-3 text-sm font-medium truncate text-slate-700">{rOutlineFile.name}</span>
                        </div>
                        <button
                          onClick={() => { setROutlineFile(null); setRStatus("idle"); }}
                          className="p-1 hover:bg-slate-200 rounded-lg transition"
                          title="Remove outline"
                        >
                          <TrashIcon />
                        </button>
                      </div>
                    ) : (
                      <div
                        onClick={() => rOutlineInputRef.current?.click()}
                        className="border-2 border-dashed border-slate-200 hover:border-slate-400 hover:bg-slate-50/50 rounded-xl p-4 text-center cursor-pointer transition"
                      >
                        <p className="text-xs font-semibold text-slate-600">+ Upload Outline File</p>
                        <p className="text-[10px] text-slate-400 mt-0.5">Supports .docx and .txt</p>
                        <input
                          ref={rOutlineInputRef}
                          type="file"
                          accept=".docx,.txt"
                          onChange={handleROutlineChange}
                          className="hidden"
                        />
                      </div>
                    )}
                    <p className="text-[11px] text-amber-600 mt-2 font-medium">
                      ⚠️ Upload a separate outline file only. Do not upload the same manuscript file here.
                    </p>
                  </div>
                </div>

                {/* Info disclaimer */}
                <div className="mt-4 p-3 bg-slate-50 rounded-xl border border-slate-100 text-xs text-slate-500">
                  Only exact lines from the uploaded or pasted outline will be converted. No automatic heading detection is used.
                </div>

                {/* Status Screen */}
                {rStatus !== "idle" && (
                  <div className={`mt-6 p-4 rounded-xl text-sm ${
                    rStatus === "processing" ? "bg-indigo-50 border border-indigo-100 text-indigo-800" :
                    rStatus === "success" ? "bg-emerald-50 border border-emerald-100 text-emerald-800" :
                    "bg-rose-50 border border-rose-100 text-rose-800"
                  }`}>
                    <div className="flex items-start">
                      {rStatus === "processing" && (
                        <svg className="animate-spin h-5 w-5 text-indigo-600 mr-2 flex-shrink-0" fill="none" viewBox="0 0 24 24">
                          <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                          <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
                        </svg>
                      )}
                      {rStatus === "success" && <CheckCircleIcon />}
                      {rStatus === "error" && <ErrorIcon />}
                      <div>
                        <p className="font-semibold">{rStatus === "processing" ? "Processing..." : rStatus === "success" ? "Completed" : "Error"}</p>
                        <p className="mt-0.5 text-xs opacity-90">{rMessage}</p>
                      </div>
                    </div>

                    {rStatus === "success" && (
                      <div className="mt-3 pt-3 border-t border-slate-200/50 text-xs space-y-3 text-slate-600">
                        <div className="grid grid-cols-2 gap-3 bg-slate-100/60 p-3 rounded-lg border border-slate-200/50 font-medium text-slate-700">
                          <div className="col-span-2"><span className="text-slate-500">Main DOCX file name:</span> <span className="font-bold text-slate-900">{rMainFile?.name}</span></div>
                          <div className="col-span-2"><span className="text-slate-500">Outline source file name or pasted outline:</span> <span className="font-bold text-slate-900">{rOutlineSource}</span></div>
                          <div><span className="text-slate-500">Total outline targets extracted:</span> <span className="font-bold text-slate-900">{rTotalOutlineTargets}</span></div>
                          <div><span className="text-slate-500">Total paragraphs converted by exact outline match:</span> <span className="font-bold text-slate-900">{rConvertedCount}</span></div>
                          <div><span className="text-slate-500">Total automatic detections used:</span> <span className="font-bold text-slate-900">0</span></div>
                          <div><span className="text-slate-500">Skipped H1 matches:</span> <span className="font-bold text-slate-900">{rSkippedH1Count}</span></div>
                          <div><span className="text-slate-500">Original file size:</span> <span className="font-bold text-slate-900">{formatSize(rOriginalSize)}</span></div>
                          <div><span className="text-slate-500">Output file size:</span> <span className="font-bold text-slate-900">{formatSize(rOutputSize)}</span></div>
                        </div>

                        {rOutputSize > 2 * rOriginalSize && (
                          <div className="bg-rose-50 border border-rose-200 text-rose-800 rounded-lg p-2.5 font-semibold text-xs mt-2 flex items-center">
                            <ErrorIcon /> Warning: Output file size increased unexpectedly. Please check compression settings.
                          </div>
                        )}

                        {/* Pre-save / processing report */}
                        <div className="border border-slate-200 rounded-xl p-3 bg-white space-y-3">
                          <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Processing Report</h4>
                          
                          <div className="grid grid-cols-2 gap-2 text-[11px]">
                            <div>Total outline targets found: <span className="font-bold text-slate-800">{rTotalOutlineTargets}</span></div>
                            <div>Total paragraphs converted: <span className="font-bold text-slate-800">{rConvertedCount}</span></div>
                          </div>

                          {/* List of converted paragraph texts */}
                          {rConvertedParagraphs.length > 0 && (
                            <div className="pt-2 border-t border-slate-100">
                              <span className="font-semibold text-slate-700 block mb-1">Converted paragraph list ({rConvertedParagraphs.length}):</span>
                              <div className="max-h-[100px] overflow-y-auto font-mono text-[10px] text-slate-600 space-y-1 bg-slate-50 p-2 rounded-lg border border-slate-100">
                                {rConvertedParagraphs.map((text, idx) => (
                                  <div key={idx}>• {text}</div>
                                ))}
                              </div>
                            </div>
                          )}

                          {/* List of outline targets not found */}
                          {rUnmatchedOutlineTargets.length > 0 && (
                            <div className="pt-2 border-t border-slate-100">
                              <span className="font-semibold text-amber-800 block mb-1">Outline targets not found ({rUnmatchedOutlineTargets.length}):</span>
                              <div className="max-h-[100px] overflow-y-auto font-mono text-[10px] text-amber-700 space-y-1 bg-amber-50/50 p-2 rounded-lg border border-amber-100">
                                {rUnmatchedOutlineTargets.map((text, idx) => (
                                  <div key={idx}>• {text}</div>
                                ))}
                              </div>
                            </div>
                          )}

                          {/* List of outline lines skipped because they are H1/title/frontmatter */}
                          {rSkippedOutlineTargets.length > 0 && (
                            <div className="pt-2 border-t border-slate-100">
                              <span className="font-semibold text-slate-500 block mb-1">Outline lines skipped because they are H1/title/frontmatter ({rSkippedOutlineTargets.length}):</span>
                              <div className="max-h-[100px] overflow-y-auto font-mono text-[10px] text-slate-500 space-y-1 bg-slate-100 p-2 rounded-lg border border-slate-200">
                                {rSkippedOutlineTargets.map((text, idx) => (
                                  <div key={idx}>• {text}</div>
                                ))}
                              </div>
                            </div>
                          )}

                          {/* List of skipped ambiguous matches */}
                          {rSkippedAmbiguousMatches.length > 0 && (
                            <div className="pt-2 border-t border-slate-100">
                              <span className="font-semibold text-rose-800 block mb-1">Skipped Ambiguous Matches ({rSkippedAmbiguousMatches.length}):</span>
                              <div className="max-h-[100px] overflow-y-auto font-mono text-[10px] text-rose-700 space-y-1 bg-rose-50/30 p-2 rounded-lg border border-rose-100">
                                {rSkippedAmbiguousMatches.map((item, idx) => (
                                  <div key={idx}>
                                    • <span className="font-semibold">{item.text}</span> <span className="italic text-[9px] text-rose-500">({item.reason})</span>
                                  </div>
                                ))}
                              </div>
                            </div>
                          )}
                        </div>

                        <div className="text-emerald-700 font-semibold flex items-center pt-1 border-t border-slate-200/40">
                          ✓ No text was rewritten. Recipe names and chapter subtopics were formatted as Heading 2 with bold OFF.
                        </div>

                        {rConvertedCount === 0 && rMainFile && (
                          <button
                            onClick={() => triggerDownload(rMainFile, rMainFile.name)}
                            className="mt-3 w-full bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2 px-3 rounded-lg border border-slate-200 transition text-center"
                          >
                            Download unchanged file
                          </button>
                        )}
                      </div>
                    )}

                    {rWarnings.length > 0 && (
                      <div className="mt-2 text-xs text-amber-700 space-y-1 bg-amber-50 p-2 rounded-lg border border-amber-100">
                        <div className="font-semibold flex items-center"><ExclamationIcon /> Warnings:</div>
                        {rWarnings.map((w, idx) => (
                          <div key={idx} className="pl-6">• {w}</div>
                        ))}
                      </div>
                    )}
                  </div>
                )}
              </div>

              <div className="mt-8">
                <button
                  onClick={runRecipeH2}
                  disabled={
                    rStatus === "processing" ||
                    !(
                      rMainFile &&
                      (rOutlineFile || rOutlineText.trim().length > 0) &&
                      !isSameFile(rMainFile, rOutlineFile)
                    )
                  }
                  className="w-full bg-indigo-600 hover:bg-indigo-700 disabled:bg-slate-200 disabled:text-slate-400 text-white font-bold py-3 px-4 rounded-xl transition shadow-sm hover:shadow active:scale-[0.99] flex items-center justify-center"
                >
                  Format Recipe Book H2
                </button>
              </div>
            </div>

          </div>
        </section>


        {/* ====================================================
            SECTION 2: IMAGE EXTRACTOR
            ==================================================== */}
        <section className="space-y-6">
          <div className="border-b border-slate-200 pb-3 flex items-center justify-between">
            <h2 className="text-2xl font-bold text-slate-900 flex items-center">
              <span className="w-2.5 h-6 bg-indigo-600 rounded-full mr-3 block"></span>
              Image Extractor
            </h2>
            <span className="text-xs text-slate-500 bg-slate-200/60 px-2 py-0.5 rounded font-semibold uppercase tracking-wider">
              Batch Assets
            </span>
          </div>

          <div className="bg-white rounded-2xl border border-slate-200/80 p-8 shadow-sm hover:shadow-md transition-all">
            <div className="max-w-3xl">
              <p className="text-sm text-slate-600">
                Extracts embedded image resources directly from DOCX, PPTX, XLSX, and ZIP files. Each archive is searched, images are converted to JPG in-browser, and a consolidated ZIP folder structure is returned.
              </p>
            </div>

            {/* Drag & Drop Upload Zone */}
            <div className="mt-8 grid grid-cols-1 lg:grid-cols-3 gap-8">
              
              <div className="lg:col-span-2 space-y-4">
                <div
                  onClick={() => ieInputRef.current?.click()}
                  className="border-2 border-dashed border-slate-200 hover:border-indigo-400 hover:bg-slate-50/50 rounded-2xl p-10 text-center cursor-pointer transition flex flex-col items-center justify-center min-h-[220px]"
                >
                  <FileZipIcon />
                  <p className="mt-3 text-base font-bold text-slate-700">Click to upload multiple archives</p>
                  <p className="text-xs text-slate-400 mt-1">Supports .docx, .pptx, .xlsx, .zip</p>
                  <input
                    ref={ieInputRef}
                    type="file"
                    multiple
                    accept=".docx,.pptx,.xlsx,.zip"
                    onChange={handleIeFilesChange}
                    className="hidden"
                  />
                </div>
              </div>

              {/* Uploaded File List */}
              <div className="border border-slate-200/80 rounded-2xl p-4 bg-slate-50/50 flex flex-col justify-between min-h-[220px]">
                <div>
                  <h4 className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-3">
                    Selected Files ({ieFiles.length})
                  </h4>
                  {ieFiles.length === 0 ? (
                    <p className="text-xs text-slate-400 italic">No files selected. Upload files to see them here.</p>
                  ) : (
                    <div className="max-h-[160px] overflow-y-auto space-y-2 pr-1">
                      {ieFiles.map((file, idx) => {
                        const fileExt = file.name.split(".").pop()?.toLowerCase();
                        return (
                          <div key={idx} className="flex items-center justify-between bg-white border border-slate-100 rounded-lg p-2 text-xs">
                            <span className="font-semibold truncate max-w-[150px] text-slate-700" title={file.name}>
                              {file.name}
                            </span>
                            <div className="flex items-center gap-2">
                              <span className="text-[10px] text-slate-400 font-medium bg-slate-100 px-1.5 py-0.5 rounded">
                                {fileExt?.toUpperCase()}
                              </span>
                              <button
                                onClick={() => removeIeFile(idx)}
                                className="p-1 hover:bg-slate-100 rounded text-slate-400 hover:text-rose-500"
                                title="Remove file"
                              >
                                ✕
                              </button>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
                {ieFiles.length > 0 && (
                  <button
                    onClick={() => { setIeFiles([]); setIeStatus("idle"); setIeMessage(""); }}
                    className="w-full text-center mt-3 text-xs text-rose-600 hover:text-rose-700 font-semibold transition"
                  >
                    Clear All Files
                  </button>
                )}
              </div>

            </div>

            {/* Status Panel / Live Processing Logs */}
            {ieStatus !== "idle" && (
              <div className={`mt-8 p-6 rounded-2xl text-sm border ${
                ieStatus === "processing" ? "bg-slate-50 border-slate-200 text-slate-800" :
                ieStatus === "success" ? "bg-emerald-50 border-emerald-100 text-emerald-800" :
                "bg-rose-50 border-rose-100 text-rose-800"
              }`}>
                <div className="flex items-start">
                  {ieStatus === "processing" && (
                    <svg className="animate-spin h-5 w-5 text-indigo-600 mr-3 flex-shrink-0" fill="none" viewBox="0 0 24 24">
                      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
                    </svg>
                  )}
                  {ieStatus === "success" && <CheckCircleIcon />}
                  {ieStatus === "error" && <ErrorIcon />}
                  <div>
                    <h5 className="font-bold">{ieStatus === "processing" ? "Processing Archives..." : ieStatus === "success" ? "Completed" : "Extraction Failed"}</h5>
                    <p className="mt-0.5 text-xs opacity-90">{ieMessage}</p>
                  </div>
                </div>

                {/* Console Logs */}
                {ieLogs.length > 0 && (
                  <div className="mt-4 p-4 bg-slate-900 text-slate-100 rounded-xl font-mono text-[10px] space-y-1.5 max-h-[180px] overflow-y-auto">
                    {ieLogs.map((log, idx) => (
                      <div key={idx} className={log.includes("completed") ? "text-emerald-400" : log.includes("Converting") ? "text-blue-300" : "text-slate-300"}>
                        {log}
                      </div>
                    ))}
                  </div>
                )}

                {/* Warnings */}
                {ieWarnings.length > 0 && (
                  <div className="mt-3 text-xs text-amber-700 space-y-1 bg-amber-50 p-2.5 rounded-lg border border-amber-100 font-semibold">
                    <div className="flex items-center font-bold"><ExclamationIcon /> Warnings:</div>
                    {ieWarnings.map((w, idx) => (
                      <div key={idx} className="pl-6 font-medium">• {w}</div>
                    ))}
                  </div>
                )}
              </div>
            )}

            <div className="mt-8 border-t border-slate-100 pt-6">
              <button
                onClick={runImageExtraction}
                disabled={ieStatus === "processing" || ieFiles.length === 0}
                className="bg-indigo-600 hover:bg-indigo-700 disabled:bg-slate-200 disabled:text-slate-400 text-white font-bold py-3.5 px-6 rounded-xl transition shadow-sm hover:shadow active:scale-[0.99] flex items-center justify-center text-sm"
              >
                <DownloadIcon />
                Extract Images and Download ZIP
              </button>
            </div>
          </div>
        </section>

      </main>

      <footer className="max-w-6xl mx-auto px-6 mt-16 text-center text-xs text-slate-400 border-t border-slate-200/60 pt-8">
        <p>© 2026 LEAPS Document Tools. Powered by Next.js client-side utilities.</p>
        <p className="mt-1 font-medium text-slate-400/90">No upload files or user content is stored on any server.</p>
      </footer>
    </div>
  );
}
