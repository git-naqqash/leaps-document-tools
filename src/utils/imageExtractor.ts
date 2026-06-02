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

export async function extractImagesFromFiles(
  files: File[],
  onProgress: (status: string) => void
): Promise<{ blob: Blob; warnings: string[] }> {
  const masterZip = new JSZip();
  const warnings: string[] = [];

  for (let fileIndex = 0; fileIndex < files.length; fileIndex++) {
    const file = files[fileIndex];
    onProgress(`Processing file ${fileIndex + 1} of ${files.length}: ${file.name}...`);

    const extension = file.name.split(".").pop()?.toLowerCase() || "";
    const folderName = file.name.replace(/\s+/g, "_") + "_extracted";

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
        }
      });

      if (imageFiles.length === 0) {
        warnings.push(`No images found in ${file.name}.`);
      } else {
        const fileFolder = masterZip.folder(folderName);
        if (!fileFolder) {
          throw new Error(`Failed to create directory in zip for ${file.name}`);
        }

        let savedCount = 0;

        // Extract and convert each image
        for (let imgIdx = 0; imgIdx < imageFiles.length; imgIdx++) {
          const { relativePath, fileObj } = imageFiles[imgIdx];
          
          onProgress(`File ${fileIndex + 1}/${files.length} (${file.name}): Converting image ${imgIdx + 1}/${imageFiles.length}...`);

          try {
            const imgBlob = await fileObj.async("blob");
            const convertedBlob = await convertImageToJpg(imgBlob);

            const imageNumber = String(savedCount + 1).padStart(3, "0");
            const outFileName = `image-${imageNumber}.jpg`;
            fileFolder.file(outFileName, convertedBlob);
            savedCount++;
          } catch (error: unknown) {
            const errMsg = error instanceof Error ? error.message : String(error);
            warnings.push(`Failed to convert image ${relativePath} in ${file.name} to JPG: ${errMsg}`);
          }
        }
      }
    } catch (zipError: unknown) {
      const zipErrMsg = zipError instanceof Error ? zipError.message : String(zipError);
      warnings.push(`Archive reading error in ${file.name}: ${zipErrMsg}`);
    }
  }

  onProgress("Packaging master ZIP archive...");
  const finalZipBlob = await masterZip.generateAsync({ type: "blob" });
  onProgress("Extraction completed successfully!");
  return { blob: finalZipBlob, warnings };
}
