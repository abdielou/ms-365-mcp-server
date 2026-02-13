import { readFileSync, statSync } from "fs";
import path from "path";
const MIME_TYPES = {
  // Microsoft Office formats
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ".doc": "application/msword",
  ".xls": "application/vnd.ms-excel",
  ".ppt": "application/vnd.ms-powerpoint",
  // Documents
  ".pdf": "application/pdf",
  ".txt": "text/plain",
  ".csv": "text/csv",
  ".json": "application/json",
  ".xml": "application/xml",
  // Images
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".png": "image/png",
  ".gif": "image/gif",
  ".bmp": "image/bmp",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  // Audio
  ".mp3": "audio/mpeg",
  ".wav": "audio/wav",
  ".ogg": "audio/ogg",
  ".m4a": "audio/mp4",
  // Video
  ".mp4": "video/mp4",
  ".avi": "video/x-msvideo",
  ".mov": "video/quicktime",
  ".wmv": "video/x-ms-wmv",
  // Archives
  ".zip": "application/zip",
  ".rar": "application/x-rar-compressed",
  ".7z": "application/x-7z-compressed",
  ".tar": "application/x-tar",
  ".gz": "application/gzip",
  // Web
  ".html": "text/html",
  ".css": "text/css",
  ".js": "application/javascript",
  ".ts": "application/typescript"
};
const DEFAULT_MIME_TYPE = "application/octet-stream";
const MAX_FILE_SIZE = 4 * 1024 * 1024 * 1024;
function getMimeType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  return MIME_TYPES[ext] || DEFAULT_MIME_TYPE;
}
function readFileForUpload(filePath) {
  let stats;
  try {
    stats = statSync(filePath);
  } catch (error) {
    throw new Error(`File not found or inaccessible: ${filePath}`);
  }
  if (!stats.isFile()) {
    throw new Error(`Path is not a file: ${filePath}`);
  }
  if (stats.size > MAX_FILE_SIZE) {
    const sizeMB = (stats.size / (1024 * 1024)).toFixed(2);
    const limitMB = (MAX_FILE_SIZE / (1024 * 1024)).toFixed(2);
    throw new Error(
      `File size (${sizeMB} MB) exceeds maximum limit (${limitMB} MB)`
    );
  }
  let content;
  try {
    content = readFileSync(filePath);
  } catch (error) {
    throw new Error(`Failed to read file: ${error.message}`);
  }
  return {
    content,
    mimeType: getMimeType(filePath),
    size: stats.size,
    fileName: path.basename(filePath)
  };
}
function validateFilePath(filePath) {
  if (typeof filePath !== "string") {
    return false;
  }
  if (filePath.trim() === "") {
    return false;
  }
  return true;
}
export {
  getMimeType,
  readFileForUpload,
  validateFilePath
};
