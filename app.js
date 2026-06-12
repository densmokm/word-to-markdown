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
    generatedMarkdown = htmlToMarkdown(result.value).trim() + "\n";
    markdownOutput.value = generatedMarkdown;
    preview.innerHTML = result.value || "Nothing to preview yet.";
    preview.classList.remove("empty");
    downloadButton.disabled = false;
    copyButton.disabled = false;
    wordCount.textContent = `${generatedMarkdown.length.toLocaleString()} characters`;
    setStatus(result.messages.length ? `Converted with ${result.messages.length} note(s). Review complex merged cells.` : "Converted successfully. Your document never left this browser.");
  } catch (error) {
    console.error(error);
    setStatus("Conversion failed. Check that the file is a valid .docx document.");
  } finally {
    convertButton.disabled = false;
  }
});

markdownOutput.addEventListener("input", () => {
  generatedMarkdown = markdownOutput.value;
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
  } catch {
    markdownOutput.select();
    document.execCommand("copy");
  }
  setStatus("Markdown copied to your clipboard.");
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

function htmlToMarkdown(html) {
  const documentFragment = new DOMParser().parseFromString(html, "text/html");
  return Array.from(documentFragment.body.childNodes)
    .map(node => convertNode(node, 0))
    .join("\n\n")
    .replace(/\n{3,}/g, "\n\n");
}

function convertNode(node, depth) {
  if (node.nodeType === Node.TEXT_NODE) return cleanText(node.textContent);
  if (node.nodeType !== Node.ELEMENT_NODE) return "";
  const tag = node.tagName.toLowerCase();
  const content = () => Array.from(node.childNodes).map(child => convertNode(child, depth)).join("");
  if (/^h[1-6]$/.test(tag)) return `${"#".repeat(Number(tag[1]))} ${inlineContent(node)}`;
  if (tag === "p") {
    const text = plainText(node);
    return looksLikeCode(text) ? fencedCodeBlock(text) : inlineContent(node);
  }
  if (tag === "pre") return fencedCodeBlock(node.textContent || "");
  if (tag === "strong" || tag === "b") return `**${content()}**`;
  if (tag === "em" || tag === "i") return `*${content()}*`;
  if (tag === "u") return content();
  if (tag === "code") return `\`${content()}\``;
  if (tag === "br") return "<br>";
  if (tag === "a") return `[${inlineContent(node)}](${node.getAttribute("href") || ""})`;
  if (tag === "ul" || tag === "ol") return convertList(node, depth + 1, tag === "ol");
  if (tag === "table") return convertTable(node);
  if (tag === "img") return node.getAttribute("alt") ? `![${escapeMarkdown(node.getAttribute("alt"))}]()` : "![Image]()";
  if (tag === "blockquote") return content().split("\n").map(line => `> ${line}`).join("\n");
  return content();
}

function convertList(list, depth, ordered) {
  return Array.from(list.children)
    .filter(child => child.tagName.toLowerCase() === "li")
    .map((item, index) => {
      const marker = ordered ? `${index + 1}.` : "-";
      const indent = "  ".repeat(Math.max(0, depth - 1));
      const direct = Array.from(item.childNodes)
        .filter(node => !(node.nodeType === Node.ELEMENT_NODE && ["ul", "ol"].includes(node.tagName.toLowerCase())))
        .map(node => convertNode(node, depth))
        .join("")
        .trim();
      const nested = Array.from(item.children)
        .filter(child => ["ul", "ol"].includes(child.tagName.toLowerCase()))
        .map(child => "\n" + convertList(child, depth + 1, child.tagName.toLowerCase() === "ol"))
        .join("");
      return `${indent}${marker} ${direct}${nested}`;
    })
    .join("\n");
}

function convertTable(table) {
  const rows = Array.from(table.rows);
  if (!rows.length) return "";
  const matrix = rows.map(row => Array.from(row.cells).map(cell => tableCellContent(cell)));
  const columnCount = Math.max(...matrix.map(row => row.length));
  const normalisedRows = matrix.map(row => [...row, ...Array(columnCount - row.length).fill("")]);
  const header = normalisedRows[0];
  const separator = Array(columnCount).fill("---");
  return [header, separator, ...normalisedRows.slice(1)]
    .map(row => `| ${row.join(" | ")} |`)
    .join("\n");
}

function tableCellContent(cell) {
  const text = plainText(cell);
  if (looksLikeCode(text)) return `<pre><code>${escapeHtml(text)}</code></pre>`;
  return Array.from(cell.childNodes)
    .map(node => convertNode(node, 0))
    .join(" ")
    .replace(/\n+/g, "<br>")
    .replace(/\|/g, "\\|")
    .replace(/\s+/g, " ")
    .trim();
}

function inlineContent(node) {
  return Array.from(node.childNodes)
    .map(child => convertNode(child, 0))
    .join("")
    .replace(/\s+/g, " ")
    .trim();
}

function plainText(node) {
  return (node.textContent || "").replace(/\u00a0/g, " ").trim();
}

function cleanText(value) {
  return value.replace(/\u00a0/g, " ").replace(/\s+/g, " ");
}

function looksLikeCode(value) {
  const text = value.trim();
  if (!text) return false;
  const sqlKeywords = /\b(SELECT|INSERT|UPDATE|DELETE|CREATE|ALTER|DROP|FROM|WHERE|JOIN|RETURNS|LANGUAGE|DECLARE|BEGIN|END|EXCEPTION|PERFORM|RAISE|COALESCE|TIMESTAMP|UUID|FUNCTION)\b/i;
  const codeSignals = /(--|;|\$\$|:=|\bINTO\b|\bNULL\b|\bCOUNT\s*\()/i;
  return text.length > 80 && sqlKeywords.test(text) && codeSignals.test(text);
}

function fencedCodeBlock(value) {
  const code = value.replace(/\u00a0/g, " ").trim();
  return `\`\`\`sql\n${code}\n\`\`\``;
}

function escapeHtml(value) {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeMarkdown(value) {
  return value.replace(/[\\`*_{}\[\]()#+\-.!|]/g, "\\$&");
}
