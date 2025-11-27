import { slides_v1 } from 'googleapis';
import { UpdateSlideNotesArgs } from '../schemas.js';
import { handleGoogleApiError } from '../utils/errorHandler.js';

/**
 * Finds the speaker notes shape using the speakerNotesObjectId from NotesProperties
 * Per Google Slides API best practices: https://developers.google.com/workspace/slides/api/guides/notes
 */
const findNotesShape = (
  notesPage: slides_v1.Schema$Page | undefined
): { shapeId: string; hasExistingText: boolean; textLength: number } | null => {
  if (!notesPage?.pageElements) return null;

  // Use the speakerNotesObjectId to find the correct shape
  const speakerNotesObjectId = notesPage.notesProperties?.speakerNotesObjectId;
  if (!speakerNotesObjectId) {
    return null;
  }

  // Find the shape with the matching object ID
  for (const element of notesPage.pageElements) {
    if (element.objectId === speakerNotesObjectId && element.shape?.text) {
      const textContent =
        element.shape.text.textElements?.map((te) => te.textRun?.content || '').join('') || '';
      const hasText = textContent.trim().length > 0;
      return {
        shapeId: element.objectId,
        hasExistingText: hasText,
        textLength: textContent.length,
      };
    }
  }

  // Shape doesn't exist yet - API will auto-create it when we perform a text operation
  // Return the speakerNotesObjectId anyway so operations can proceed
  return {
    shapeId: speakerNotesObjectId,
    hasExistingText: false,
    textLength: 0,
  };
};

/**
 * Updates speaker notes for a specific slide.
 * @param slides The authenticated Google Slides API client.
 * @param args The arguments for updating slide notes.
 * @returns A promise resolving to the MCP response content.
 * @throws McpError if the Google API call fails.
 */
export const updateSlideNotesTool = async (slides: slides_v1.Slides, args: UpdateSlideNotesArgs) => {
  try {
    // First, get the page to find the notes shape
    const pageResponse = await slides.presentations.pages.get({
      presentationId: args.presentationId,
      pageObjectId: args.pageObjectId,
    });

    const notesPage = pageResponse.data.slideProperties?.notesPage;
    if (!notesPage) {
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                success: false,
                error: 'No notes page found for this slide',
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

    const notesShape = findNotesShape(notesPage);
    if (!notesShape) {
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                success: false,
                error:
                  'No speakerNotesObjectId found in notes page properties. The notes page may not be properly initialized.',
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

    // Build batch update requests based on operation
    const requests: slides_v1.Schema$Request[] = [];
    const operation = args.operation || 'replace';

    switch (operation) {
      case 'replace':
        // Delete all existing text and insert new text
        if (notesShape.hasExistingText) {
          requests.push({
            deleteText: {
              objectId: notesShape.shapeId,
              textRange: { type: 'ALL' },
            },
          });
        }
        if (args.notes.trim()) {
          requests.push({
            insertText: {
              objectId: notesShape.shapeId,
              text: args.notes,
              insertionIndex: 0,
            },
          });
        }
        break;

      case 'append':
        // Insert text at the end
        if (args.notes.trim()) {
          requests.push({
            insertText: {
              objectId: notesShape.shapeId,
              text: args.notes,
              insertionIndex: notesShape.textLength,
            },
          });
        }
        break;

      case 'insert':
        // Insert text at specific index
        if (args.insertionIndex === undefined) {
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(
                  {
                    success: false,
                    error: 'insertionIndex is required for insert operation',
                  },
                  null,
                  2
                ),
              },
            ],
            isError: true,
          };
        }
        if (args.notes.trim()) {
          requests.push({
            insertText: {
              objectId: notesShape.shapeId,
              text: args.notes,
              insertionIndex: args.insertionIndex,
            },
          });
        }
        break;

      case 'delete':
        // Delete text in range
        if (args.deleteStartIndex === undefined || args.deleteEndIndex === undefined) {
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(
                  {
                    success: false,
                    error: 'deleteStartIndex and deleteEndIndex are required for delete operation',
                  },
                  null,
                  2
                ),
              },
            ],
            isError: true,
          };
        }
        requests.push({
          deleteText: {
            objectId: notesShape.shapeId,
            textRange: {
              type: 'FIXED_RANGE',
              startIndex: args.deleteStartIndex,
              endIndex: args.deleteEndIndex,
            },
          },
        });
        break;
    }

    // Execute the batch update
    if (requests.length > 0) {
      await slides.presentations.batchUpdate({
        presentationId: args.presentationId,
        requestBody: {
          requests,
        },
      });
    }

    const result = {
      success: true,
      presentationId: args.presentationId,
      pageObjectId: args.pageObjectId,
      notesShapeId: notesShape.shapeId,
      operation,
      message: `Speaker notes ${operation}d successfully`,
    };

    return {
      content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
    };
  } catch (error: unknown) {
    throw handleGoogleApiError(error, 'update_slide_notes');
  }
};
