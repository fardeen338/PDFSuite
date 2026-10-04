import { $, formatBytes, safeFileName, downloadBlob, setupDragDrop } from "./utils.js";

export function initWordToPdf() {
  const container = $("tool-word-to-pdf");
  if (!container) return;

  const dropZone = container.querySelector(".drop-zone");
  const fileInput = container.querySelector(".file-input");
  const browseBtn = container.querySelector(".browse-btn");
  const filePanel = container.querySelector(".file-panel");
  const fileName = container.querySelector(".file-name");
  const fileMeta = container.querySelector(".file-meta");
  const removeBtn = container.querySelector(".remove-btn");
  const convertBtn = container.querySelector(".convert-btn");
  const previewBox = container.querySelector(".word-preview-box");
  const message = container.querySelector(".message");

  let selectedFile = null;
  let htmlContent = "";

  function showMessage(text, type = "error") {
    message.textContent = text;
    message.className = `message ${type}`;
    message.classList.remove("hidden");
  }

  function clearMessage() {
    message.textContent = "";
    message.className = "message hidden";
  }

  function reset() {
    selectedFile = null;
    htmlContent = "";
    fileInput.value = "";
    filePanel.classList.add("hidden");
    previewBox.innerHTML = "";
    previewBox.classList.add("hidden");
    convertBtn.disabled = true;
    clearMessage();
  }

  async function selectFile(file) {
    reset();
    if (!file) return;

    if (!file.name.toLowerCase().endsWith(".docx")) {
      showMessage("Please upload a Microsoft Word document (.docx).");
      return;
    }

    selectedFile = file;
    fileName.textContent = file.name;
    fileMeta.textContent = formatBytes(file.size);
    filePanel.classList.remove("hidden");

    try {
      if (!window.mammoth) {
        throw new Error("Mammoth.js library not loaded from CDN.");
      }

      const arrayBuffer = await file.arrayBuffer();
      const result = await window.mammoth.convertToHtml({ arrayBuffer });
      htmlContent = result.value;

      if (!htmlContent) {
        showMessage("No readable text found in this Word file.");
        return;
      }

      previewBox.innerHTML = htmlContent;
      previewBox.classList.remove("hidden");
      convertBtn.disabled = false;
      showMessage("Word document parsed. Click 'Convert to PDF' below to generate your PDF.", "success");
    } catch (e) {
      console.error(e);
      showMessage(`Could not read Word file: ${e.message}`);
    }
  }

  async function convertToPdf() {
    if (!htmlContent || !selectedFile || !window.html2pdf) return;
    clearMessage();
    convertBtn.disabled = true;
    showMessage("Generating vector PDF from document...", "success");

    try {
      const opt = {
        margin: [15, 15, 15, 15],
        filename: `${safeFileName(selectedFile.name)}.pdf`,
        image: { type: "jpeg", quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true },
        jsPDF: { unit: "mm", format: "a4", orientation: "portrait" }
      };

      await window.html2pdf().set(opt).from(previewBox).save();
      showMessage("PDF generated and downloaded successfully!", "success");
    } catch (e) {
      console.error(e);
      showMessage(`Conversion failed: ${e.message}`);
    } finally {
      convertBtn.disabled = false;
    }
  }

  browseBtn.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", (e) => selectFile(e.target.files[0]));
  removeBtn.addEventListener("click", reset);
  convertBtn.addEventListener("click", convertToPdf);
  setupDragDrop(dropZone, selectFile);
}
