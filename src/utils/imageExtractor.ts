import JSZip from "jszip";

// Helper to convert any browser-supported image Blob to JPG using HTML5 Canvas
export async function convertImageToJpg(blob: Blob): Promise<Blob> {
  return new Promise((resolve, reject) => {
    // Create an object URL for the blob
    const url = URL.createObjectURL(blob);
    const img = new Image();

    img.onload = () => {
      // Clean up object URL
      URL.revokeObjectURL(url);

      const canvas = document.createElement("canvas");
      // Use natural sizes to preserve full resolution
      canvas.width = img.naturalWidth || img.width;
      canvas.height = img.naturalHeight || img.height;

      const ctx = canvas.getContext("2d");
      if (!ctx) {
        reject(new Error("Could not get 2d context from canvas"));
        return;
      }

      // Draw a solid white background (crucial for transparent PNGs/WebPs to avoid black backgrounds in JPG)
      ctx.fillStyle = "#FFFFFF";
      ctx.fillRect(0, 0, canvas.width, canvas.height);

      // Draw the image
      ctx.drawImage(img, 0, 0);

      // Export as jpeg with high quality (0.9)
      canvas.toBlob(
        (resultBlob) => {
          if (resultBlob) {
            resolve(resultBlob);
          } else {
            reject(new Error("Canvas export returned null blob"));
          }
        },
        "image/jpeg",
        0.9
      );
    };

    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Browser failed to load image resource."));
    };

    img.src = url;
  });
}

interface ExtractionResult {
  fileName: string;
  totalFound: number;
  totalConverted: number;
  failures: string[];
  unsupported: string[];
}

export async function extractImagesFromFiles(
  files: File[],
  onProgress: (status: string) => void
): Promise<Blob> {
  const masterZip = new JSZip();

  for (let fileIndex = 0; fileIndex < files.length; fileIndex++) {
    const file = files[fileIndex];
    onProgress(`Processing file ${fileIndex + 1} of ${files.length}: ${file.name}...`);

    const extension = file.name.split(".").pop()?.toLowerCase() || "";
    const folderName = file.name.replace(/\s+/g, "_") + "_extracted";
    const fileFolder = masterZip.folder(folderName);

    if (!fileFolder) {
      throw new Error(`Failed to create directory in zip for ${file.name}`);
    }

    const result: ExtractionResult = {
      fileName: file.name,
      totalFound: 0,
      totalConverted: 0,
      failures: [],
      unsupported: [],
    };

    try {
      const arrayBuffer = await file.arrayBuffer();
      const zip = await JSZip.loadAsync(arrayBuffer);

      // Determine directories to scan based on file type
      let isOfficeFile = true;
      let targetPrefix = "";

      if (extension === "docx") {
        targetPrefix = "word/media/";
      } else if (extension === "pptx") {
        targetPrefix = "ppt/media/";
      } else if (extension === "xlsx") {
        targetPrefix = "xl/media/";
      } else if (extension === "zip") {
        isOfficeFile = false;
      } else {
        throw new Error(`Unsupported master file type: ${extension}`);
      }

      // Collect all candidate image files from the archive
      const imageFiles: { relativePath: string; fileObj: JSZip.JSZipObject }[] = [];
      zip.forEach((relativePath, fileObj) => {
        if (fileObj.dir) return;

        const fileExt = relativePath.split(".").pop()?.toLowerCase() || "";
        const isImageExtension = ["jpg", "jpeg", "png", "webp", "gif"].includes(fileExt);

        if (isImageExtension) {
          if (isOfficeFile) {
            if (relativePath.startsWith(targetPrefix)) {
              imageFiles.push({ relativePath, fileObj });
            }
          } else {
            // For standard ZIPs, scan all directories
            imageFiles.push({ relativePath, fileObj });
          }
        } else {
          // If we are scanning media folders and find non-images or general zip scanning
          if (!isOfficeFile) {
            result.unsupported.push(`${relativePath} (Not an image format)`);
          }
        }
      });

      result.totalFound = imageFiles.length;

      // Extract and convert each image
      for (let imgIdx = 0; imgIdx < imageFiles.length; imgIdx++) {
        const { relativePath, fileObj } = imageFiles[imgIdx];
        const fileExt = relativePath.split(".").pop()?.toLowerCase() || "";
        const imageNumber = String(imgIdx + 1).padStart(3, "0");
        const outFileName = `image-${imageNumber}.jpg`;

        onProgress(`File ${fileIndex + 1}/${files.length} (${file.name}): Converting image ${imgIdx + 1}/${imageFiles.length}...`);

        try {
          const imgBlob = await fileObj.async("blob");
          const convertedBlob = await convertImageToJpg(imgBlob);

          fileFolder.file(outFileName, convertedBlob);
          result.totalConverted++;
        } catch (error: unknown) {
          // Fallback: save original format if canvas conversion fails
          const fallbackName = `image-${imageNumber}.${fileExt}`;
          const errMsg = error instanceof Error ? error.message : String(error);
          try {
            const originalBlob = await fileObj.async("blob");
            fileFolder.file(fallbackName, originalBlob);
            result.failures.push(
              `${relativePath} - Conversion failed (${errMsg}). Copied original file as ${fallbackName}.`
            );
          } catch (readError: unknown) {
            const readErrMsg = readError instanceof Error ? readError.message : String(readError);
            result.failures.push(
              `${relativePath} - Critical read failure: ${readErrMsg}`
            );
          }
        }
      }
    } catch (zipError: unknown) {
      const zipErrMsg = zipError instanceof Error ? zipError.message : String(zipError);
      result.failures.push(`Archive reading error: ${zipErrMsg}`);
    }

    // Write report.txt inside the folder
    const reportContent = generateReportText(result);
    fileFolder.file("report.txt", reportContent);
  }

  onProgress("Packaging master ZIP archive...");
  const finalZipBlob = await masterZip.generateAsync({ type: "blob" });
  onProgress("Extraction completed successfully!");
  return finalZipBlob;
}

function generateReportText(result: ExtractionResult): string {
  let text = `====================================================\n`;
  text += `IMAGE EXTRACTION REPORT\n`;
  text += `====================================================\n`;
  text += `Source Filename: ${result.fileName}\n`;
  text += `Total Images Found: ${result.totalFound}\n`;
  text += `Total Images Converted to JPG: ${result.totalConverted}\n\n`;

  if (result.failures.length > 0) {
    text += `Failed Conversions (copied in original format):\n`;
    result.failures.forEach((fail) => {
      text += `- ${fail}\n`;
    });
    text += `\n`;
  } else {
    text += `Failed Conversions: None\n\n`;
  }

  if (result.unsupported.length > 0) {
    text += `Ignored Archive Files (Non-image files in ZIP):\n`;
    result.unsupported.forEach((file) => {
      text += `- ${file}\n`;
    });
  } else {
    text += `Ignored Archive Files: None\n`;
  }

  text += `====================================================\n`;
  return text;
}
