import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';

export interface ImageDownloadResult {
  filePath: string;
  mimeType: string;
  size: number;
}

export interface ImageDownloadOptions {
  outputDir?: string;
  filename?: string;
}

/**
 * Detects image format from HTTP Content-Type header or magic bytes
 */
export const detectImageFormat = (
  contentType: string | null,
  buffer: Buffer
): { extension: string; mimeType: string } => {
  // Check magic bytes first (more reliable)
  if (buffer.length >= 8) {
    // PNG: 89 50 4E 47 0D 0A 1A 0A
    if (buffer[0] === 0x89 && buffer[1] === 0x50 && buffer[2] === 0x4e && buffer[3] === 0x47) {
      return { extension: 'png', mimeType: 'image/png' };
    }
    // JPEG: FF D8 FF
    if (buffer[0] === 0xff && buffer[1] === 0xd8 && buffer[2] === 0xff) {
      return { extension: 'jpg', mimeType: 'image/jpeg' };
    }
    // GIF: 47 49 46 38
    if (buffer[0] === 0x47 && buffer[1] === 0x49 && buffer[2] === 0x46 && buffer[3] === 0x38) {
      return { extension: 'gif', mimeType: 'image/gif' };
    }
    // WebP: RIFF....WEBP
    if (
      buffer[0] === 0x52 &&
      buffer[1] === 0x49 &&
      buffer[2] === 0x46 &&
      buffer[3] === 0x46 &&
      buffer[8] === 0x57 &&
      buffer[9] === 0x45 &&
      buffer[10] === 0x42 &&
      buffer[11] === 0x50
    ) {
      return { extension: 'webp', mimeType: 'image/webp' };
    }
  }

  // Fall back to Content-Type header
  if (contentType) {
    const mimeMap: Record<string, { extension: string; mimeType: string }> = {
      'image/png': { extension: 'png', mimeType: 'image/png' },
      'image/jpeg': { extension: 'jpg', mimeType: 'image/jpeg' },
      'image/gif': { extension: 'gif', mimeType: 'image/gif' },
      'image/webp': { extension: 'webp', mimeType: 'image/webp' },
    };
    const match = mimeMap[contentType.split(';')[0].trim().toLowerCase()];
    if (match) return match;
  }

  // Default to PNG if unknown
  return { extension: 'png', mimeType: 'image/png' };
};

/**
 * Creates a unique temp directory for storing downloaded images
 */
export const createTempImageDir = (presentationId: string, pageId?: string): string => {
  const basePath = path.join(os.tmpdir(), 'slides-images', presentationId);
  const targetPath = pageId ? path.join(basePath, pageId) : basePath;
  fs.mkdirSync(targetPath, { recursive: true });
  return targetPath;
};

/**
 * Downloads an image from a URL using native Node.js fetch
 */
export const downloadImageFromUrl = async (
  url: string,
  options: ImageDownloadOptions = {}
): Promise<ImageDownloadResult> => {
  const response = await fetch(url);

  if (!response.ok) {
    throw new Error(`Failed to download image: HTTP ${response.status} ${response.statusText}`);
  }

  const arrayBuffer = await response.arrayBuffer();
  const buffer = Buffer.from(arrayBuffer);

  if (buffer.length === 0) {
    throw new Error('Downloaded image is empty');
  }

  const contentType = response.headers.get('content-type');
  const { extension, mimeType } = detectImageFormat(contentType, buffer);

  const outputDir = options.outputDir || os.tmpdir();
  const filename = options.filename ? `${options.filename}.${extension}` : `image_${Date.now()}.${extension}`;

  const filePath = path.join(outputDir, filename);

  fs.writeFileSync(filePath, buffer);

  return {
    filePath,
    mimeType,
    size: buffer.length,
  };
};
