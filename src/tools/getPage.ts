import { slides_v1 } from 'googleapis';
import { GetPageArgs } from '../schemas.js';
import { handleGoogleApiError } from '../utils/errorHandler.js';
import { filterPageElements } from '../utils/elementFilter.js';

/**
 * Gets details about a specific page (slide) in a presentation.
 * @param slides The authenticated Google Slides API client.
 * @param args The arguments for getting the page.
 * @returns A promise resolving to the MCP response content.
 * @throws McpError if the Google API call fails.
 */
export const getPageTool = async (slides: slides_v1.Slides, args: GetPageArgs) => {
  try {
    const response = await slides.presentations.pages.get({
      presentationId: args.presentationId,
      pageObjectId: args.pageObjectId,
    });

    let pageData = response.data;

    // Apply filtering by default (can be disabled with filterSmallElements: false)
    const shouldFilter = args.filterSmallElements !== false;
    if (shouldFilter && pageData.pageElements) {
      const originalCount = pageData.pageElements.length;
      const filteredElements = filterPageElements(pageData.pageElements);
      const filteredCount = originalCount - filteredElements.length;

      pageData = {
        ...pageData,
        pageElements: filteredElements,
        // Add metadata about filtering
        ...(filteredCount > 0 && {
          filteringApplied: {
            enabled: true,
            elementsFiltered: filteredCount,
            originalCount,
            finalCount: filteredElements.length,
            note: 'Small decorative elements (< 100 points with no text) filtered to reduce context',
          } as any,
        }),
      };
    }

    return {
      content: [{ type: 'text', text: JSON.stringify(pageData, null, 2) }],
    };
  } catch (error: unknown) {
    throw handleGoogleApiError(error, 'get_page');
  }
};
