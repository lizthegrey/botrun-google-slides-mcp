import { slides_v1 } from 'googleapis';
import { ListSlideImagesArgs } from '../schemas.js';
import { handleGoogleApiError } from '../utils/errorHandler.js';

export interface ImageInfo {
  elementId: string;
  contentUrl: string;
  sourceUrl?: string;
  width?: number;
  height?: number;
  description?: string;
  parentGroupId?: string;
}

/**
 * Recursively extracts image information from page elements,
 * including images nested within groups.
 */
const extractImagesFromElements = (
  elements: slides_v1.Schema$PageElement[] | undefined,
  parentGroupId?: string
): ImageInfo[] => {
  if (!elements) return [];

  const images: ImageInfo[] = [];

  for (const element of elements) {
    const elementId = element.objectId;
    if (!elementId) continue;

    // Handle direct image elements
    if (element.image) {
      const image = element.image;
      if (image.contentUrl) {
        images.push({
          elementId,
          contentUrl: image.contentUrl,
          sourceUrl: image.sourceUrl || undefined,
          width: element.size?.width?.magnitude ?? undefined,
          height: element.size?.height?.magnitude ?? undefined,
          description: element.description || undefined,
          parentGroupId,
        });
      }
    }

    // Handle shapes with image fills
    const shapeBackgroundFill = element.shape?.shapeProperties?.shapeBackgroundFill as any;
    if (shapeBackgroundFill?.stretchedPictureFill?.contentUrl) {
      images.push({
        elementId,
        contentUrl: shapeBackgroundFill.stretchedPictureFill.contentUrl,
        width: element.size?.width?.magnitude ?? undefined,
        height: element.size?.height?.magnitude ?? undefined,
        description: element.description || 'Shape with image fill',
        parentGroupId,
      });
    }

    // Recursively handle groups
    if (element.elementGroup?.children) {
      images.push(...extractImagesFromElements(element.elementGroup.children, elementId));
    }
  }

  return images;
};

/**
 * Lists all images on a slide with their metadata.
 * @param slides The authenticated Google Slides API client.
 * @param args The arguments for listing slide images.
 * @returns A promise resolving to the MCP response content.
 * @throws McpError if the Google API call fails.
 */
export const listSlideImagesTool = async (slides: slides_v1.Slides, args: ListSlideImagesArgs) => {
  try {
    const response = await slides.presentations.pages.get({
      presentationId: args.presentationId,
      pageObjectId: args.pageObjectId,
    });

    const page = response.data;
    const images = extractImagesFromElements(page.pageElements);

    const result = {
      presentationId: args.presentationId,
      pageObjectId: args.pageObjectId,
      imageCount: images.length,
      images,
    };

    return {
      content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
    };
  } catch (error: unknown) {
    throw handleGoogleApiError(error, 'list_slide_images');
  }
};
