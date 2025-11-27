import { slides_v1 } from 'googleapis';
import { DownloadImageArgs } from '../schemas.js';
import { downloadImageFromUrl, createTempImageDir } from '../utils/imageDownloader.js';
import { handleGoogleApiError } from '../utils/errorHandler.js';

/**
 * Finds an image element on a page by its object ID
 */
const findImageElement = (
  elements: slides_v1.Schema$PageElement[] | undefined,
  targetId: string
): string | null => {
  if (!elements) return null;

  for (const element of elements) {
    // Direct image
    if (element.objectId === targetId && element.image?.contentUrl) {
      return element.image.contentUrl;
    }

    // Image in group
    if (element.elementGroup?.children) {
      const found = findImageElement(element.elementGroup.children, targetId);
      if (found) return found;
    }
  }

  return null;
};

/**
 * Downloads a slide image by element ID.
 * Fetches the image URL from the slide and validates it exists.
 * Returns the file path so Claude's Read() tool can process it.
 * @param slides The authenticated Google Slides API client.
 * @param args The arguments for downloading the image.
 * @returns A promise resolving to the MCP response content.
 */
export const downloadImageTool = async (slides: slides_v1.Slides, args: DownloadImageArgs) => {
  try {
    // Fetch the page to find the image element by its ID
    const pageResponse = await slides.presentations.pages.get({
      presentationId: args.presentationId,
      pageObjectId: args.pageObjectId,
    });

    const imageUrl = findImageElement(pageResponse.data.pageElements, args.elementId);
    if (!imageUrl) {
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                success: false,
                error: `Image element ${args.elementId} not found on slide ${args.pageObjectId}`,
                presentationId: args.presentationId,
                pageObjectId: args.pageObjectId,
                elementId: args.elementId,
              },
              null,
              2
            ),
          },
        ],
        isError: true,
      };
    }

    // Create output directory in temp location
    const outputDir = args.outputDir || createTempImageDir(args.presentationId, args.pageObjectId);

    const result = await downloadImageFromUrl(imageUrl, {
      outputDir,
      filename: args.filename || args.elementId,
    });

    const response = {
      success: true,
      filePath: result.filePath,
      mimeType: result.mimeType,
      size: result.size,
      elementId: args.elementId,
      message: `Image downloaded successfully. Use Read("${result.filePath}") to view it.`,
    };

    return {
      content: [{ type: 'text', text: JSON.stringify(response, null, 2) }],
    };
  } catch (error: unknown) {
    throw handleGoogleApiError(error, 'download_image');
  }
};
