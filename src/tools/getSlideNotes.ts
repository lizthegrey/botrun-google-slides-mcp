import { slides_v1 } from 'googleapis';
import { GetSlideNotesArgs } from '../schemas.js';
import { handleGoogleApiError } from '../utils/errorHandler.js';

/**
 * Extracts text from a specific shape (the speaker notes shape)
 */
const extractNotesText = (
  elements: slides_v1.Schema$PageElement[] | undefined,
  speakerNotesObjectId: string | null | undefined
): string => {
  if (!elements || !speakerNotesObjectId) return '';

  // Find the specific speaker notes shape by its object ID
  for (const element of elements) {
    if (element.objectId === speakerNotesObjectId && element.shape?.text?.textElements) {
      return element.shape.text.textElements.map((te) => te.textRun?.content || '').join('');
    }
  }

  return '';
};

/**
 * Gets speaker notes for a specific slide.
 * @param slides The authenticated Google Slides API client.
 * @param args The arguments for getting slide notes.
 * @returns A promise resolving to the MCP response content.
 * @throws McpError if the Google API call fails.
 */
export const getSlideNotesTool = async (slides: slides_v1.Slides, args: GetSlideNotesArgs) => {
  try {
    const response = await slides.presentations.pages.get({
      presentationId: args.presentationId,
      pageObjectId: args.pageObjectId,
    });

    const page = response.data;
    const notesPage = page.slideProperties?.notesPage;

    if (!notesPage) {
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                presentationId: args.presentationId,
                pageObjectId: args.pageObjectId,
                notes: '',
                message: 'No notes page found for this slide',
              },
              null,
              2
            ),
          },
        ],
      };
    }

    // Use the speakerNotesObjectId to identify the correct shape
    const speakerNotesObjectId = notesPage.notesProperties?.speakerNotesObjectId;
    const notesText = extractNotesText(notesPage.pageElements, speakerNotesObjectId);

    const result = {
      presentationId: args.presentationId,
      pageObjectId: args.pageObjectId,
      notes: notesText,
      notesPageObjectId: notesPage.objectId,
      speakerNotesObjectId: speakerNotesObjectId || null,
    };

    return {
      content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
    };
  } catch (error: unknown) {
    throw handleGoogleApiError(error, 'get_slide_notes');
  }
};
