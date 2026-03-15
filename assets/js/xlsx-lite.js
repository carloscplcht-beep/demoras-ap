(function () {
  "use strict";

  const XML_MIME_TYPE = "application/xml";
  const UTF8_DECODER = new TextDecoder("utf-8");

  function readUInt16(view, offset) {
    return view.getUint16(offset, true);
  }

  function readUInt32(view, offset) {
    return view.getUint32(offset, true);
  }

  function decodeBytes(bytes) {
    return UTF8_DECODER.decode(bytes);
  }

  function normalizePath(path) {
    return String(path || "").replace(/\\/g, "/").replace(/^\//, "");
  }

  function joinPath(baseFile, target) {
    if (!target) {
      return "";
    }

    const normalizedTarget = normalizePath(target);
    if (/^[a-z]+:\/\//i.test(normalizedTarget)) {
      return normalizedTarget;
    }

    if (normalizedTarget.startsWith("xl/")) {
      return normalizedTarget;
    }

    const basePath = normalizePath(baseFile);
    const baseParts = basePath.split("/");
    baseParts.pop();

    const targetParts = normalizedTarget.split("/");
    const resolved = baseParts.concat(targetParts);
    const finalParts = [];

    for (const part of resolved) {
      if (!part || part === ".") {
        continue;
      }
      if (part === "..") {
        finalParts.pop();
      } else {
        finalParts.push(part);
      }
    }

    return finalParts.join("/");
  }

  function getAttributeByLocalName(node, localName) {
    if (!node || !node.attributes) {
      return "";
    }

    for (const attribute of Array.from(node.attributes)) {
      if (attribute.localName === localName || attribute.name === localName) {
        return attribute.value;
      }
    }

    return "";
  }

  function getElements(node, localName) {
    if (!node || !node.getElementsByTagNameNS) {
      return [];
    }
    return Array.from(node.getElementsByTagNameNS("*", localName));
  }

  function firstElement(node, localName) {
    return getElements(node, localName)[0] || null;
  }

  function parseXml(xmlText) {
    const documentNode = new DOMParser().parseFromString(xmlText, XML_MIME_TYPE);
    const parserError = firstElement(documentNode, "parsererror");
    if (parserError) {
      throw new Error("No se ha podido interpretar el XML interno del archivo Excel.");
    }
    return documentNode;
  }

  function locateEndOfCentralDirectory(bytes) {
    const minSize = 22;
    for (let index = bytes.length - minSize; index >= 0; index -= 1) {
      if (
        bytes[index] === 0x50 &&
        bytes[index + 1] === 0x4b &&
        bytes[index + 2] === 0x05 &&
        bytes[index + 3] === 0x06
      ) {
        return index;
      }
    }

    throw new Error("El archivo no parece ser un libro XLSX valido.");
  }

  async function inflateRaw(bytes) {
    if (typeof DecompressionStream === "undefined") {
      throw new Error(
        "Este navegador no soporta la descompresion local necesaria para leer archivos XLSX. Utilice una version reciente de Edge o Chrome."
      );
    }

    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function unzipXlsx(arrayBuffer) {
    const bytes = new Uint8Array(arrayBuffer);
    const view = new DataView(arrayBuffer);
    const eocdOffset = locateEndOfCentralDirectory(bytes);
    const totalEntries = readUInt16(view, eocdOffset + 10);
    const centralDirectoryOffset = readUInt32(view, eocdOffset + 16);
    const entries = new Map();

    let cursor = centralDirectoryOffset;

    for (let index = 0; index < totalEntries; index += 1) {
      const signature = readUInt32(view, cursor);
      if (signature !== 0x02014b50) {
        throw new Error("La estructura ZIP interna del libro no es valida.");
      }

      const generalPurposeBitFlag = readUInt16(view, cursor + 8);
      const compressionMethod = readUInt16(view, cursor + 10);
      const compressedSize = readUInt32(view, cursor + 20);
      const fileNameLength = readUInt16(view, cursor + 28);
      const extraFieldLength = readUInt16(view, cursor + 30);
      const fileCommentLength = readUInt16(view, cursor + 32);
      const localHeaderOffset = readUInt32(view, cursor + 42);
      const fileNameBytes = bytes.slice(cursor + 46, cursor + 46 + fileNameLength);
      const fileName = normalizePath(decodeBytes(fileNameBytes));

      const localHeaderSignature = readUInt32(view, localHeaderOffset);
      if (localHeaderSignature !== 0x04034b50) {
        throw new Error("No se ha encontrado una entrada local valida en el libro XLSX.");
      }

      const localFileNameLength = readUInt16(view, localHeaderOffset + 26);
      const localExtraFieldLength = readUInt16(view, localHeaderOffset + 28);
      const dataOffset = localHeaderOffset + 30 + localFileNameLength + localExtraFieldLength;
      const compressedBytes = bytes.slice(dataOffset, dataOffset + compressedSize);

      let uncompressedBytes;
      if (compressionMethod === 0) {
        uncompressedBytes = compressedBytes;
      } else if (compressionMethod === 8) {
        uncompressedBytes = await inflateRaw(compressedBytes);
      } else {
        throw new Error("El metodo de compresion del libro XLSX no es compatible.");
      }

      entries.set(fileName, {
        name: fileName,
        flags: generalPurposeBitFlag,
        bytes: uncompressedBytes
      });

      cursor += 46 + fileNameLength + extraFieldLength + fileCommentLength;
    }

    return entries;
  }

  function getEntryText(entries, path) {
    const entry = entries.get(normalizePath(path));
    return entry ? decodeBytes(entry.bytes) : "";
  }

  function parseSharedStrings(entries) {
    const sharedStringsXml = getEntryText(entries, "xl/sharedStrings.xml");
    if (!sharedStringsXml) {
      return [];
    }

    const sharedStringsDocument = parseXml(sharedStringsXml);
    const stringItems = getElements(sharedStringsDocument, "si");

    return stringItems.map((item) => {
      const textNodes = getElements(item, "t");
      return textNodes.map((node) => node.textContent || "").join("");
    });
  }

  function parseWorkbookRelationships(entries, workbookPath) {
    const normalizedWorkbookPath = normalizePath(workbookPath);
    const workbookFileName = normalizedWorkbookPath.split("/").pop();
    const workbookDir = normalizedWorkbookPath.split("/");
    workbookDir.pop();
    const relsPath = workbookDir.concat(["_rels", workbookFileName + ".rels"]).join("/");
    const relationshipsXml = getEntryText(entries, relsPath);
    if (!relationshipsXml) {
      return new Map();
    }

    const relationshipsDocument = parseXml(relationshipsXml);
    const relationships = getElements(relationshipsDocument, "Relationship");
    const relationshipMap = new Map();

    relationships.forEach((relationship) => {
      const id = getAttributeByLocalName(relationship, "Id");
      const target = getAttributeByLocalName(relationship, "Target");
      if (id && target) {
        relationshipMap.set(id, joinPath(workbookPath, target));
      }
    });

    return relationshipMap;
  }

  function parseWorkbook(entries) {
    const workbookPath = "xl/workbook.xml";
    const workbookXml = getEntryText(entries, workbookPath);
    if (!workbookXml) {
      throw new Error("No se ha encontrado el libro interno del archivo XLSX.");
    }

    const workbookDocument = parseXml(workbookXml);
    const relationshipMap = parseWorkbookRelationships(entries, workbookPath);
    const sheetNodes = getElements(workbookDocument, "sheet");
    const sheets = sheetNodes.map((sheetNode) => {
      const name = getAttributeByLocalName(sheetNode, "name");
      const relationId = getAttributeByLocalName(sheetNode, "id");
      return {
        name,
        path: relationshipMap.get(relationId) || ""
      };
    });

    if (!sheets.length || !sheets[0].path) {
      throw new Error("No se ha podido localizar la primera hoja del libro.");
    }

    return sheets[0];
  }

  function columnReferenceToIndex(reference) {
    const columnReference = String(reference || "").replace(/\d+/g, "").toUpperCase();
    let index = 0;

    for (let position = 0; position < columnReference.length; position += 1) {
      index = index * 26 + (columnReference.charCodeAt(position) - 64);
    }

    return index > 0 ? index - 1 : 0;
  }

  function readCellValue(cellNode, sharedStrings) {
    const type = getAttributeByLocalName(cellNode, "t");

    if (type === "inlineStr") {
      return getElements(cellNode, "t")
        .map((node) => node.textContent || "")
        .join("");
    }

    const valueNode = firstElement(cellNode, "v");
    if (!valueNode) {
      return "";
    }

    const rawValue = valueNode.textContent || "";

    if (type === "s") {
      return sharedStrings[Number(rawValue)] || "";
    }

    if (type === "b") {
      return rawValue === "1" ? "TRUE" : "FALSE";
    }

    return rawValue;
  }

  function parseSheet(entries, sheetPath, sharedStrings) {
    const sheetXml = getEntryText(entries, sheetPath);
    if (!sheetXml) {
      throw new Error("No se ha podido leer la primera hoja del archivo Excel.");
    }

    const sheetDocument = parseXml(sheetXml);
    const rowNodes = getElements(sheetDocument, "row");
    const rawRows = [];
    let maxColumnLength = 0;

    rowNodes.forEach((rowNode) => {
      const values = [];
      let maxColumnIndex = -1;
      let sequentialColumnIndex = 0;
      const cellNodes = Array.from(rowNode.childNodes).filter(
        (childNode) => childNode.nodeType === 1 && childNode.localName === "c"
      );

      cellNodes.forEach((cellNode) => {
        const reference = getAttributeByLocalName(cellNode, "r");
        const columnIndex = reference ? columnReferenceToIndex(reference) : sequentialColumnIndex;
        values[columnIndex] = readCellValue(cellNode, sharedStrings);
        maxColumnIndex = Math.max(maxColumnIndex, columnIndex);
        sequentialColumnIndex = columnIndex + 1;
      });

      const rowHasContent = values.some((value) => String(value || "").trim() !== "");
      if (rowHasContent) {
        const denseRow = Array.from({ length: maxColumnIndex + 1 }, (_, index) =>
          values[index] == null ? "" : values[index]
        );
        rawRows.push(denseRow);
        maxColumnLength = Math.max(maxColumnLength, denseRow.length);
      }
    });

    return rawRows.map((row) =>
      Array.from({ length: maxColumnLength }, (_, index) => (row[index] == null ? "" : row[index]))
    );
  }

  async function readWorkbook(file) {
    if (!file) {
      throw new Error("No se ha seleccionado ningun archivo.");
    }

    const buffer = await file.arrayBuffer();
    const entries = await unzipXlsx(buffer);
    const firstSheet = parseWorkbook(entries);
    const sharedStrings = parseSharedStrings(entries);
    const rawRows = parseSheet(entries, firstSheet.path, sharedStrings);

    return {
      workbookName: file.name,
      sheetName: firstSheet.name || "Hoja 1",
      rawRows
    };
  }

  window.XLSXLite = {
    readWorkbook
  };
})();
