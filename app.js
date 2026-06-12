const fileInput = document.getElementById("fileInput");
const dropZone = document.getElementById("dropZone");
const convertButton = document.getElementById("convertButton");
const downloadButton = document.getElementById("downloadButton");
const copyButton = document.getElementById("copyButton");
const clearButton = document.getElementById("clearButton");
const markdownOutput = document.getElementById("markdownOutput");
const preview = document.getElementById("preview");
const status = document.getElementById("status");
const wordCount = document.getElementById("wordCount");

let selectedFile = null;
let generatedMarkdown = "";

function setStatus(message) {
  status.textContent = message;
}

function setFile(file) {
  if (!file) return;

  if (!file.name.toLowerCase().endsWith(".docx")) {
    setStatus("Please choose a .docx Word document.");
    return;
  }

  selectedFile = file;
  convertButton.disabled = false;
  clearButton.disabled = false;
  setStatus(`Ready to convert: ${file.name}`);
}

fileInput.addEventListener("change", (event) => {
  setFile(event.target.files[0]);
});

["dragenter", "dragover"].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.add("dragging");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.remove("dragging");
  });
});

dropZone.addEventListener("drop", (event) => {
  setFile(event.dataTransfer.files[0]);
});

convertButton.addEventListener("click", async () => {
  if (!selectedFile) return;

  try {
    convertButton.disabled = true;
    setStatus("Converting document locally in your browser...");

    const arrayBuffer = await selectedFile.arrayBuffer();
    const result = await mammoth.convertToMarkdown({ arrayBuffer });

    generatedMarkdown = result.value.trim() + "\n";
    markdownOutput.value = generatedMarkdown;
    preview.innerHTML = markdownToSafePreview(generatedMarkdown);
    preview.classList.remove("empty");

    downloadButton.disabled = false;
    copyButton.disabled = false;
    wordCount.textContent = `${generatedMarkdown.length.toLocaleString()} characters`;

    if (result.messages.length) {
      setStatus(`Converted with ${result.messages.length} note(s). Complex Word formatting may need a quick tidy.`);
    } else {
      setStatus("Converted successfully. Your document never left this browser.");
    }
  } catch (error) {
    console.error(error);
    setStatus("Conversion failed. Check that the file is a valid .docx document.");
  } finally {
    convertButton.disabled = false;
  }
});

markdownOutput.addEventListener("input", () => {
  generatedMarkdown = markdownOutput.value;
  preview.innerHTML = markdownToSafePreview(generatedMarkdown);
  preview.classList.toggle("empty", !generatedMarkdown.trim());
  wordCount.textContent = `${generatedMarkdown.length.toLocaleString()} characters`;
  downloadButton.disabled = !generatedMarkdown.trim();
  copyButton.disabled = !generatedMarkdown.trim();
});

downloadButton.addEventListener("click", () => {
  if (!generatedMarkdown) return;

  const blob = new Blob([generatedMarkdown], { type: "text/markdown;charset=utf-8" });
  const link = document.createElement("a");
  const baseName = selectedFile ? selectedFile.name.replace(/\.docx$/i, "") : "converted-document";

  link.href = URL.createObjectURL(blob);
  link.download = `${baseName}.md`;
  link.click();
  URL.revokeObjectURL(link.href);
});

copyButton.addEventListener("click", async () => {
  if (!generatedMarkdown) return;

  try {
    await navigator.clipboard.writeText(generatedMarkdown);
    setStatus("Markdown copied to your clipboard.");
  } catch {
    markdownOutput.select();
    document.execCommand("copy");
    setStatus("Markdown copied to your clipboard.");
  }
});

clearButton.addEventListener("click", () => {
  selectedFile = null;
  generatedMarkdown = "";
  fileInput.value = "";
  markdownOutput.value = "";
  preview.textContent = "Nothing to preview yet.";
  preview.classList.add("empty");
  convertButton.disabled = true;
  downloadButton.disabled = true;
  copyButton.disabled = true;
  clearButton.disabled = true;
  wordCount.textContent = "0 characters";
  setStatus("Choose a Word document to begin.");
});

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function markdownToSafePreview(markdown) {
  if (!markdown.trim()) return "Nothing to preview yet.";

  const escaped = escapeHtml(markdown);

  return escaped
    .replace(/^###### (.*)$/gm, "<h6>$1</h6>")
    .replace(/^##### (.*)$/gm, "<h5>$1</h5>")
    .replace(/^#### (.*)$/gm, "<h4>$1</h4>")
    .replace(/^### (.*)$/gm, "<h3>$1</h3>")
    .replace(/^## (.*)$/gm, "<h2>$1</h2>")
    .replace(/^# (.*)$/gm, "<h1>$1</h1>")
    .replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>")
    .replace(/\*(.*?)\*/g, "<em>$1</em>")
    .replace(/`([^`]+)`/g, "<code>$1</code>")
    .replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, '<a href="$2" target="_blank" rel="noopener noreferrer">$1</a>')
    .replace(/\n{2,}/g, "</p><p>")
    .replace(/\n/g, "<br>")
    .replace(/^/, "<p>")
    .replace(/$/, "</p>");
}
