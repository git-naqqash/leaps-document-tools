# LEAPS Document Tools

A professional, production-ready, client-side web application designed to safely format Word document headings and extract media assets from Office archives entirely in the browser.

## Key Features

### 1. H2 Maker
Automates heading formatting for documents with no existing numbered headings. Includes two separate specialized modes:
- **Non-Recipe Book H2 Maker**: Designed for general nonfiction or fiction books. Formats clear section or subtopic titles.
- **Recipe Book H2 Maker**: Built for cookbooks. Targets recipe titles and chapter subtopics without stripping serial numbers or changing numbering format.
- **Strict Format Restrictions**:
  - Sets style level to `Heading 2`.
  - Turns Bold properties `OFF` (setting `w:val="false"` in the XML).
  - **Zero content modification**: Never rewrites, rephrases, deletes, or inserts text, serial numbers, symbols, spaces, or layout markers.
  - Skips any paragraphs styled as `Heading 1`.
  - Supports **Optional Outline Files** (.docx and .txt) to restrict updates strictly to matching outline topics.

### 2. Image Extractor
Extracts media files from standard Office containers and folder structures.
- **Supported input formats**: `.docx`, `.pptx`, `.xlsx`, and `.zip`.
- **Conversion properties**: Extracted images (PNG, WebP, JPEG, GIF first-frame) are converted to `.jpg` format client-side using the HTML5 Canvas API.
- **Consolidated ZIP output**: Generates separate subfolders per uploaded archive containing processed JPGs (`image-001.jpg`, etc.) and a detailed conversion summary (`report.txt`). Packages all folders into a final download archive `extracted-images-master.zip`.

---

## 🔒 Privacy & Safety First

- **100% Client-Side Processing**: All parsing, XML modifications, image extractions, canvas drawing, and ZIP compression occur directly in the browser's memory.
- **No Server Communication**: No API routes, databases, or third-party servers are used. Files never leave the local user machine.

---

## Supported File Formats

- **H2 Maker Inputs**: `.docx` only.
- **Outline Inputs**: `.docx` and `.txt`.
- **Image Extractor Inputs**: `.docx`, `.pptx`, `.xlsx`, and `.zip` archives.
- **Extracted Image Formats**: `.png`, `.jpg`, `.jpeg`, `.webp`, and `.gif`.

---

## Technology Stack

- **Framework**: [Next.js](https://nextjs.org/) (App Router, TypeScript)
- **Styles**: [Tailwind CSS](https://tailwindcss.com/)
- **Zip Compression**: [JSZip](https://stuk.github.io/jszip/)
- **XML Processing**: Browser Native `DOMParser` and `XMLSerializer` (preserves OpenXML namespace declarations and node ordering).

---

## How to Run Locally

### Prerequisites
Make sure you have Node.js (v18+) and npm installed on your system.

1. **Clone or download the project** to your local workspace.
2. **Install dependencies**:
   ```bash
   npm install
   ```
3. **Start the development server**:
   ```bash
   npm run dev
   ```
4. **Open your browser** and navigate to:
   [http://localhost:3000](http://localhost:3000)

---

## How to Deploy on Vercel

This app is fully compatible with Vercel and can be deployed with standard zero-config settings.

### Option 1: Direct Vercel CLI (Recommended)
1. Install Vercel CLI:
   ```bash
   npm install -g vercel
   ```
2. Run the deployment command in the project directory:
   ```bash
   vercel
   ```
3. Deploy to production:
   ```bash
   vercel --prod
   ```

### Option 2: Git Integration via Vercel Dashboard
1. Push this codebase to a git repository (GitHub, GitLab, or Bitbucket).
2. Connect your Vercel account to the repository.
3. Import the repository. Vercel will automatically detect the Next.js setup.
4. Click **Deploy**.

---

## Limitations

- **Large Files**: Since decompression and image conversion occur in browser memory, processing files larger than 100MB may experience memory throttles depending on browser capability.
- **Animated GIFs**: The image extractor exports the first frame of animated GIFs as a static JPG. Original GIFs can be retrieved by inspecting the warning logs in the downloaded zip.
- **Complex Styles**: Custom paragraph styles in Word that do not inherit from standard Word definitions might require the standard Heading 2 Style ID (`"Heading2"`).
