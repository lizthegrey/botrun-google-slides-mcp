import { slides_v1 } from 'googleapis';
import { GetPresentationArgs } from '../schemas.js';
import { handleGoogleApiError } from '../utils/errorHandler.js';
import { filterPageElements } from '../utils/elementFilter.js';

/**
 * Gets details about a Google Slides presentation.
 * @param slides The authenticated Google Slides API client.
 * @param args The arguments for getting the presentation.
 * @returns A promise resolving to the MCP response content.
 * @throws McpError if the Google API call fails.
 */
export const getPresentationTool = async (slides: slides_v1.Slides, args: GetPresentationArgs) => {
  try {
    const response = await slides.presentations.get({
      presentationId: args.presentationId,
      fields: args.fields,
    });

    let presentationData = response.data;

    // Apply filtering by default (can be disabled with filterSmallElements: false)
    const shouldFilter = args.filterSmallElements !== false;
    if (shouldFilter && presentationData.slides) {
      let totalFiltered = 0;
      let totalOriginal = 0;

      presentationData.slides = presentationData.slides.map((slide) => {
        if (slide.pageElements) {
          const originalCount = slide.pageElements.length;
          const filteredElements = filterPageElements(slide.pageElements);
          const filteredCount = originalCount - filteredElements.length;

          totalOriginal += originalCount;
          totalFiltered += filteredCount;

          return {
            ...slide,
            pageElements: filteredElements,
          };
        }
        return slide;
      });

      // Add metadata about filtering
      if (totalFiltered > 0) {
        (presentationData as any).filteringApplied = {
          enabled: true,
          totalElementsFiltered: totalFiltered,
          totalOriginalElements: totalOriginal,
          totalFinalElements: totalOriginal - totalFiltered,
          note: 'Small decorative elements filtered across all slides to reduce context',
        };
      }
    }

    return {
      content: [{ type: 'text', text: JSON.stringify(presentationData, null, 2) }],
    };
  } catch (error: unknown) {
    throw handleGoogleApiError(error, 'get_presentation');
  }
};
