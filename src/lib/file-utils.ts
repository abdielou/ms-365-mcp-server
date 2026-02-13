import { readFileSync, statSync } from 'fs';
import path from 'path';

/**
 * MIME type mapping for common file extensions.
 * Used for auto-detecting Content-Type when uploading files to OneDrive.
 */
const MIME_TYPES: Record<string, string> = {
  // Microsoft Office formats
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.doc': 'application/msword',
  '.xls': 'application/vnd.ms-excel',
  '.ppt': 'application/vnd.ms-powerpoint',

  // Documents
  '.pdf': 'application/pdf',
  '.txt': 'text/plain',
  '.csv': 'text/csv',
  '.json': 'application/json',
  '.xml': 'application/xml',

  // Images
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.png': 'image/png',
  '.gif': 'image/gif',
  '.bmp': 'image/bmp',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',

  // Audio
  '.mp3': 'audio/mpeg',
  '.wav': 'audio/wav',
  '.ogg': 'audio/ogg',
  '.m4a': 'audio/mp4',

  // Video
  '.mp4': 'video/mp4',
  '.avi': 'video/x-msvideo',
  '.mov': 'video/quicktime',
  '.wmv': 'video/x-ms-wmv',

  // Archives
  '.zip': 'application/zip',
  '.rar': 'application/x-rar-compressed',
  '.7z': 'application/x-7z-compressed',
  '.tar': 'application/x-tar',
  '.gz': 'application/gzip',

  // Web
  '.html': 'text/html',
  '.css': 'text/css',
  '.js': 'application/javascript',
  '.ts': 'application/typescript',
};

/**
 * Default MIME type for unknown file extensions.
 */
const DEFAULT_MIME_TYPE = 'application/octet-stream';

/**
 * Maximum file size for uploads (4GB - Microsoft Graph API limit).
 */
const MAX_FILE_SIZE = 4 * 1024 * 1024 * 1024; // 4GB

/**
 * Data returned after reading a file for upload.
 */
export interface FileData {
  content: Buffer;
  mimeType: string;
  size: number;
  fileName: string;
}

/**
 * Get the MIME type for a file based on its extension.
 *
 * @param filePath - Path to the file
 * @returns MIME type string (e.g., "application/pdf", "text/plain")
 */
export function getMimeType(filePath: string): string {
  const ext = path.extname(filePath).toLowerCase();
  return MIME_TYPES[ext] || DEFAULT_MIME_TYPE;
}

/**
 * Read a file from disk for uploading to OneDrive.
 * Validates file existence, size, and reads content as Buffer.
 *
 * @param filePath - Absolute or relative path to the file
 * @returns FileData object with content, MIME type, size, and filename
 * @throws Error if file doesn't exist, is a directory, exceeds size limit, or can't be read
 */
export function readFileForUpload(filePath: string): FileData {
  // Validate file exists and get stats
  let stats;
  try {
    stats = statSync(filePath);
  } catch (error) {
    throw new Error(`File not found or inaccessible: ${filePath}`);
  }

  // Check if it's a file (not directory)
  if (!stats.isFile()) {
    throw new Error(`Path is not a file: ${filePath}`);
  }

  // Check size limit
  if (stats.size > MAX_FILE_SIZE) {
    const sizeMB = (stats.size / (1024 * 1024)).toFixed(2);
    const limitMB = (MAX_FILE_SIZE / (1024 * 1024)).toFixed(2);
    throw new Error(
      `File size (${sizeMB} MB) exceeds maximum limit (${limitMB} MB)`
    );
  }

  // Read file content
  let content: Buffer;
  try {
    content = readFileSync(filePath);
  } catch (error) {
    throw new Error(`Failed to read file: ${(error as Error).message}`);
  }

  return {
    content,
    mimeType: getMimeType(filePath),
    size: stats.size,
    fileName: path.basename(filePath),
  };
}

/**
 * Validate that a value is a non-empty string representing a file path.
 * Used as a Zod refinement validator.
 *
 * @param filePath - Value to validate
 * @returns true if valid, false otherwise
 */
export function validateFilePath(filePath: unknown): filePath is string {
  if (typeof filePath !== 'string') {
    return false;
  }

  // Basic path validation
  if (filePath.trim() === '') {
    return false;
  }

  return true;
}
