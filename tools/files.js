const { graphGetRaw } = require("../graph");
const { ok, optionalPositiveNumber, requireString, userError } = require("../response");
const mammoth = require("mammoth");
const { PDFParse } = require("pdf-parse");
const XLSX = require("xlsx");
const JSZip = require("jszip");
const { XMLParser } = require("fast-xml-parser");

const MAX_BYTES = 20 * 1024 * 1024;
const DEFAULT_MAX_CHARS = 12000;

function trimText(text, maxChars) {
  if (!text) return "";
  if (text.length <= maxChars) return text;
  return `${text.slice(0, maxChars)}\n\n...(truncated)...`;
}

async function parseDocx(buffer) {
  try {
    const result = await mammoth.extractRawText({ buffer });
    return result.value;
  } catch (err) {
    throw userError("I couldn't read the text from this Word document.", {
      statusCode: 422,
      code: "file_parse_failed",
      internalMessage: err.message,
    });
  }
}

async function parsePdf(buffer) {
  try {
    const parser = new PDFParse({ data: buffer });
    const result = await parser.getText();
    await parser.destroy();
    return result.text;
  } catch (err) {
    throw userError("I couldn't read the text from this PDF.", {
      statusCode: 422,
      code: "file_parse_failed",
      internalMessage: err.message,
    });
  }
}

function parseXlsx(buffer) {
  try {
    const workbook = XLSX.read(buffer, { type: "buffer" });
    const output = [];

    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      output.push(`### Sheet: ${sheetName}\n${rows.map((row) => row.join("\t")).join("\n")}`);
    }

    return output.join("\n\n");
  } catch (err) {
    throw userError("I couldn't read the data from this Excel file.", {
      statusCode: 422,
      code: "file_parse_failed",
      internalMessage: err.message,
    });
  }
}

async function parsePptx(buffer) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const parser = new XMLParser({
      ignoreAttributes: true,
      removeNSPrefix: true,
    });

    const slides = Object.keys(zip.files).filter((path) =>
      /^ppt\/slides\/slide\d+\.xml$/i.test(path)
    );

    const output = [];

    for (const slide of slides) {
      const xml = await zip.file(slide).async("string");
      const object = parser.parse(xml);
      const textRuns = [];

      const walk = (node) => {
        if (!node || typeof node !== "object") return;
        for (const [key, value] of Object.entries(node)) {
          if (key === "t" && typeof value === "string") textRuns.push(value);
          else if (Array.isArray(value)) value.forEach(walk);
          else if (typeof value === "object") walk(value);
        }
      };

      walk(object);

      const slideNumber = slide.match(/slide(\d+)/)?.[1] || "";
      output.push(`### Slide ${slideNumber}\n${textRuns.join(" ")}`);
    }

    return output.join("\n\n");
  } catch (err) {
    throw userError("I couldn't read the text from this PowerPoint file.", {
      statusCode: 422,
      code: "file_parse_failed",
      internalMessage: err.message,
    });
  }
}

function parsePlain(buffer) {
  return buffer.toString("utf8");
}

const tools = [
  {
    name: "read_file_content",
    description:
      "Download a SharePoint file and extract readable text from supported formats such as Word, PDF, Excel, PowerPoint, text, CSV, and JSON.",
    inputSchema: {
      type: "object",
      properties: {
        driveId: { type: "string", description: "The document library ID." },
        itemId: { type: "string", description: "The file item ID." },
        maxChars: {
          type: "number",
          description: "Optional maximum number of characters to return from the extracted text.",
        },
      },
      required: ["driveId", "itemId"],
    },
  },
];

const handlers = {
  read_file_content: async ({ accessToken, params }) => {
    const driveId = requireString(params, "driveId", "a document library ID");
    const itemId = requireString(params, "itemId", "a file item ID");
    const maxChars = optionalPositiveNumber(params?.maxChars, "maxChars", { integer: true, min: 1 });
    const limit = maxChars || DEFAULT_MAX_CHARS;

    const meta = await graphGetRaw(accessToken, `/drives/${driveId}/items/${itemId}`);
    const name = meta.name || "file";
    const size = meta.size || 0;
    const ext = (name.split(".").pop() || "").toLowerCase();
    const supportedTypes = ["docx", "pdf", "xlsx", "xls", "pptx", "txt", "md", "csv", "json"];

    if (!supportedTypes.includes(ext)) {
      throw userError(
        `This file type is not supported for text extraction yet. Supported formats are ${supportedTypes
          .map((type) => `.${type}`)
          .join(", ")}.`,
        {
          statusCode: 400,
          code: "unsupported_file_type",
        }
      );
    }

    if (size > MAX_BYTES) {
      throw userError(
        `This file is too large to read here. The limit is ${(MAX_BYTES / 1024 / 1024).toFixed(0)} MB.`,
        {
          statusCode: 400,
          code: "file_too_large",
        }
      );
    }

    const buffer = await graphGetRaw(accessToken, `/drives/${driveId}/items/${itemId}/content`, true);

    let extracted = "";
    if (ext === "docx") extracted = await parseDocx(buffer);
    else if (ext === "pdf") extracted = await parsePdf(buffer);
    else if (ext === "xlsx" || ext === "xls") extracted = parseXlsx(buffer);
    else if (ext === "pptx") extracted = await parsePptx(buffer);
    else extracted = parsePlain(buffer);

    const trimmedText = trimText(extracted, limit);

    return ok(
      `File: ${name}\nType: .${ext}\nSize: ${(size / 1024).toFixed(2)} KB\n\n${trimmedText}`,
      {
        file: {
          id: meta.id,
          name,
          extension: ext,
          size,
          webUrl: meta.webUrl,
        },
        extractedText: trimmedText,
        truncated: extracted.length > limit,
      }
    );
  },
};

module.exports = { tools, handlers };
