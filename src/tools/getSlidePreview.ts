import { slides_v1 } from 'googleapis';
import { GetSlidePreviewArgs } from '../schemas.js';
import { handleGoogleApiError } from '../utils/errorHandler.js';
import { downloadImageFromUrl, createTempImageDir } from '../utils/imageDownloader.js';

/**
 * Gets a visual preview (thumbnail) of a slide and downloads it locally.
 * @param slides The authenticated Google Slides API client.
 * @param args The arguments for getting the slide preview.
 * @returns A promise resolving to the MCP response content with file path.
 * @throws McpError if the Google API call fails.
 */
export const getSlidePreviewTool = async (slides: slides_v1.Slides, args: GetSlidePreviewArgs) => {
  try {
    // Fetch the thumbnail from Google Slides API
    const response = await slides.presentations.pages.getThumbnail({
      presentationId: args.presentationId,
      pageObjectId: args.pageObjectId,
      'thumbnailProperties.thumbnailSize': args.thumbnailSize || 'LARGE',
      'thumbnailProperties.mimeType': args.mimeType || 'PNG',
    });

    const thumbnailUrl = response.data?.contentUrl;
    if (!thumbnailUrl) {
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                success: false,
                error: 'No thumbnail URL returned from API',
                presentationId: args.presentationId,
                pageObjectId: args.pageObjectId,
              },
              null,
              2
            ),
          },
        ],
        isError: true,
      };
    }

    // Download the thumbnail to local file
    const outputDir = args.outputDir || createTempImageDir(args.presentationId, args.pageObjectId);

    const result = await downloadImageFromUrl(thumbnailUrl, {
      outputDir,
      filename: args.filename || `preview_${args.pageObjectId}`,
    });

    const responseData = {
      success: true,
      filePath: result.filePath,
      mimeType: result.mimeType,
      size: result.size,
      presentationId: args.presentationId,
      pageObjectId: args.pageObjectId,
      thumbnailSize: args.thumbnailSize || 'LARGE',
      message: `Slide preview downloaded successfully. Use Read("${result.filePath}") to view the slide layout.`,
    };

    return {
      content: [{ type: 'text', text: JSON.stringify(responseData, null, 2) }],
    };
  } catch (error: unknown) {
    throw handleGoogleApiError(error, 'get_slide_preview');
  }
};
