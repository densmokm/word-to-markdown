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
let generatedMarkup = "";

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

fileInput.addEventListener("change", event => setFile(event.target.files[0]));

["dragenter", "dragover"].forEach(eventName => {
  dropZone.addEventListener(eventName, event => {
    event.preventDefault();
    dropZone.classList.add("dragging");
  });
});

["dragleave", "drop"].forEach(eventName => {
  dropZone.addEventListener(eventName, event => {
    event.preventDefault();
    dropZone.classList.remove("dragging");
  });
});

dropZone.addEventListener("drop", event => setFile(event.dataTransfer.files[0]));

convertButton.addEventListener("click", async () => {
  if (!selectedFile) return;
  try {
    convertButton.disabled = true;
    setStatus("Converting document locally in your browser...");
    const arrayBuffer = await selectedFile.arrayBuffer();
    const result = await mammoth.convertToHtml({ arrayBuffer });
    generatedMarkup = htmlToConfluence(result.value).trim() + "\n";
    markdownOutput.value = generatedMarkup;
    preview.innerHTML = result.value || "Nothing to preview yet.";
    preview.classList.remove("empty");
    downloadButton.disabled = false;
    copyButton.disabled = false;
    wordCount.textContent = `${generatedMarkup.length.toLocaleString()} characters`;
    setStatus(result.messages.length ? `Converted with ${result.messages.length} note(s). Review complex merged cells.` : "Converted successfully. Your document never left this browser.");
  } catch (error) {
    console.error(error);
    setStatus("Conversion failed. Check that the file is a valid .docx document.");
  } finally {
    convertButton.disabled = false;
  }
});

markdownOutput.addEventListener("input", () => {
  generatedMarkup = markdownOutput.value;
  wordCount.textContent = `${generatedMarkup.length.toLocaleString()} characters`;
  downloadButton.disabled = !generatedMarkup.trim();
  copyButton.disabled = !generatedMarkup.trim();
});

downloadButton.addEventListener("click", () => {
  if (!generatedMarkup) return;
  const blob = new Blob([generatedMarkup], { type: "text/plain;charset=utf-8" });
  const link = document.createElement("a");
  const baseName = selectedFile ? selectedFile.name.replace(/\.docx$/i, "") : "converted-document";
  link.href = URL.createObjectURL(blob);
  link.download = `${baseName}-confluence-wiki.txt`;
  link.click();
  URL.revokeObjectURL(link.href);
});

copyButton.addEventListener("click", async () => {
  if (!generatedMarkup) return;
  try {
    await navigator.clipboard.writeText(generatedMarkup);
  } catch {
    markdownOutput.select();
    document.execCommand("copy");
  }
  setStatus("Confluence wiki markup copied to your clipboard.");
});

clearButton.addEventListener("click", () => {
  selectedFile = null;
  generatedMarkup = "";
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

function htmlToConfluence(html) {
  const documentFragment = new DOMParser().parseFromString(html, "text/html");
  return Array.from(documentFragment.body.childNodes).map(node => convertNode(node, 0)).join("\n\n").replace(/\n{3,}/g, "\n\n");
}

function convertNode(node, depth) {
  if (node.nodeType === Node.TEXT_NODE) return cleanText(node.textContent);
  if (node.nodeType !== Node.ELEMENT_NODE) return "";
  const tag = node.tagName.toLowerCase();
  const content = () => Array.from(node.childNodes).map(child => convertNode(child, depth)).join("");
  if (/^h[1-6]$/.test(tag)) return `${tag}. ${inlineContent(node)}`;
  if (tag === "p") return inlineContent(node);
  if (tag === "strong" || tag === "b") return `*${content()}*`;
  if (tag === "em" || tag === "i") return `_${content()}_`;
  if (tag === "u") return `+${content()}+`;
  if (tag === "code") return `{{${content()}}}`;
  if (tag === "br") return "\\\\";
  if (tag === "a") return `[${inlineContent(node)}|${node.getAttribute("href") || ""}]`;
  if (tag === "ul" || tag === "ol") return convertList(node, depth + 1, tag === "ol");
  if (tag === "table") return convertTable(node);
  if (tag === "img") return node.getAttribute("alt") ? `[Image: ${node.getAttribute("alt")}]` : "[Image]";
  if (tag === "blockquote") return `{quote}\n${content()}\n{quote}`;
  return content();
}

function convertList(list, depth, ordered) {
  const marker = ordered ? "#" : "*";
  return Array.from(list.children).filter(child => child.tagName.toLowerCase() === "li").map(item => {
    const direct = Array.from(item.childNodes).filter(node => !(node.nodeType === Node.ELEMENT_NODE && ["ul", "ol"].includes(node.tagName.toLowerCase()))).map(node => convertNode(node, depth)).join("");
    const nested = Array.from(item.children).filter(child => ["ul", "ol"].includes(child.tagName.toLowerCase())).map(child => "\n" + convertList(child, depth + 1, child.tagName.toLowerCase() === "ol")).join("");
    return `${marker.repeat(depth)} ${direct.trim()}${nested}`;
  }).join("\n");
}

function convertTable(table) {
  const rows = Array.from(table.rows);
  if (!rows.length) return "";
  return rows.map((row, index) => {
    const cells = Array.from(row.cells);
    const isHeader = cells.every(cell => cell.tagName.toLowerCase() === "th") || (index === 0 && cells.some(cell => cell.tagName.toLowerCase() === "th"));
    const separator = isHeader ? "||" : "|";
    const values = cells.map(cell => tableCellContent(cell));
    return `${separator}${values.join(separator)}${separator}`;
  }).join("\n");
}

function tableCellContent(cell) {
  return Array.from(cell.childNodes).map(node => convertNode(node, 0)).join(" ")
    .replace(/\n+/g, "\\\\")
    .replace(/\|/g, "\\|")
    .replace(/\s+/g, " ")
    .trim();
}

function inlineContent(node) {
  return Array.from(node.childNodes).map(child => convertNode(child, 0)).join("").replace(/\s+/g, " ").trim();
}

function cleanText(value) {
  return value.replace(/\u00a0/g, " ").replace(/\s+/g, " ");
}
