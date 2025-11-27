import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { CallToolRequestSchema, ErrorCode, ListToolsRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import { slides_v1 } from 'googleapis';
import {
  CreatePresentationArgsSchema,
  GetPresentationArgsSchema,
  BatchUpdatePresentationArgsSchema,
  GetPageArgsSchema,
  SummarizePresentationArgsSchema,
  movePresentationSchema,
  ListSlideImagesArgsSchema,
  DownloadImageArgsSchema,
  GetSlideNotesArgsSchema,
  UpdateSlideNotesArgsSchema,
  GetSlidePreviewArgsSchema,
} from './schemas.js';
import { createPresentationTool } from './tools/createPresentation.js';
import { getPresentationTool } from './tools/getPresentation.js';
import { batchUpdatePresentationTool } from './tools/batchUpdatePresentation.js';
import { getPageTool } from './tools/getPage.js';
import { summarizePresentationTool } from './tools/summarizePresentation.js';
import { MovePresentationTool } from './tools/movePresentation.js';
import { listSlideImagesTool } from './tools/listSlideImages.js';
import { downloadImageTool } from './tools/downloadImage.js';
import { getSlideNotesTool } from './tools/getSlideNotes.js';
import { updateSlideNotesTool } from './tools/updateSlideNotes.js';
import { getSlidePreviewTool } from './tools/getSlidePreview.js';
import { executeTool } from './utils/toolExecutor.js';

export const setupToolHandlers = (server: Server, slides: slides_v1.Slides, auth?: any) => {
  server.setRequestHandler(ListToolsRequestSchema, async () => ({
    tools: [
      {
        name: 'create_presentation',
        description: 'Create a new Google Slides presentation',
        inputSchema: {
          type: 'object',
          properties: {
            title: {
              type: 'string',
              description: 'The title of the presentation.',
            },
          },
          required: ['title'],
        },
      },
      {
        name: 'get_presentation',
        description:
          'Get details about a Google Slides presentation. Filtering is ON by default to save context - set filterSmallElements=false only if you need full unfiltered details.',
        inputSchema: {
          type: 'object',
          properties: {
            presentationId: {
              type: 'string',
              description: 'The ID of the presentation to retrieve.',
            },
            fields: {
              type: 'string',
              description:
                'Optional. A mask specifying which fields to include in the response (e.g., "slides,pageSize").',
            },
            filterSmallElements: {
              type: 'boolean',
              description:
                'Optional (default: true). Filters out small decorative elements and semi-transparent shapes across all slides to save context. Set to false only if you need complete unfiltered presentation details.',
            },
          },
          required: ['presentationId'],
        },
      },
      {
        name: 'batch_update_presentation',
        description: 'Apply a batch of updates to a Google Slides presentation',
        inputSchema: {
          type: 'object',
          properties: {
            presentationId: {
              type: 'string',
              description: 'The ID of the presentation to update.',
            },
            requests: {
              type: 'array',
              description:
                'A list of update requests to apply. See Google Slides API documentation for request structures.',
              items: { type: 'object' },
            },
            writeControl: {
              type: 'object',
              description: 'Optional. Provides control over how write requests are executed.',
              properties: {
                requiredRevisionId: { type: 'string' },
                targetRevisionId: { type: 'string' },
              },
            },
          },
          required: ['presentationId', 'requests'],
        },
      },
      {
        name: 'get_page',
        description:
          'Get details about a specific page (slide) in a presentation. Filtering is ON by default to save context - set filterSmallElements=false only if you need full unfiltered details.',
        inputSchema: {
          type: 'object',
          properties: {
            presentationId: {
              type: 'string',
              description: 'The ID of the presentation.',
            },
            pageObjectId: {
              type: 'string',
              description: 'The object ID of the page (slide) to retrieve. NOTE: If copied from a URL like "slide=id.g123abc", remove the "id." prefix - use just "g123abc".',
            },
            filterSmallElements: {
              type: 'boolean',
              description:
                'Optional (default: true). Filters out small decorative elements and semi-transparent shapes to save context. Set to false only if you need complete unfiltered slide details.',
            },
          },
          required: ['presentationId', 'pageObjectId'],
        },
      },
      {
        name: 'summarize_presentation',
        description: 'Extract text content from all slides in a presentation for summarization purposes',
        inputSchema: {
          type: 'object',
          properties: {
            presentationId: {
              type: 'string',
              description: 'The ID of the presentation to summarize.',
            },
            include_notes: {
              type: 'boolean',
              description: 'Optional. Whether to include speaker notes in the summary (default: false).',
            },
          },
          required: ['presentationId'],
        },
      },
      {
        name: 'move_presentation',
        description: 'Move or copy a presentation to a specific Google Drive folder',
        inputSchema: {
          type: 'object',
          properties: {
            presentationId: {
              type: 'string',
              description: 'The ID of the presentation to move.',
            },
            folderId: {
              type: 'string',
              description: 'The ID of the target Google Drive folder.',
            },
            copyInstead: {
              type: 'boolean',
              description: 'If true, creates a copy in the target folder instead of moving (default: false).',
            },
            newName: {
              type: 'string',
              description: 'New name for the presentation (only used when copyInstead is true).',
            },
          },
          required: ['presentationId', 'folderId'],
        },
      },
      {
        name: 'list_slide_images',
        description:
          'List all images on a slide with their URLs and metadata. Use this first to see what images exist on a slide, then use download_image to fetch specific images you want to analyze. Returns contentUrl for each image which can be passed to download_image.',
        inputSchema: {
          type: 'object',
          properties: {
            presentationId: {
              type: 'string',
              description: 'The ID of the presentation.',
            },
            pageObjectId: {
              type: 'string',
              description: 'The object ID of the page (slide) to extract images from. NOTE: If copied from a URL like "slide=id.g123abc", remove the "id." prefix - use just "g123abc".',
            },
          },
          required: ['presentationId', 'pageObjectId'],
        },
      },
      {
        name: 'download_image',
        description:
          'Download a slide image to local file by element ID. Provide presentationId + pageObjectId + elementId (from list_slide_images). Tool auto-fetches and validates the image URL from the slide. After downloading, use Read(filePath) to view. Automatically detects PNG/JPG format.',
        inputSchema: {
          type: 'object',
          properties: {
            presentationId: {
              type: 'string',
              description: 'The presentation ID.',
            },
            pageObjectId: {
              type: 'string',
              description: 'The slide ID. NOTE: Remove "id." prefix from URLs.',
            },
            elementId: {
              type: 'string',
              description: 'The image element ID from list_slide_images. Tool will fetch and validate the URL from the slide.',
            },
            filename: {
              type: 'string',
              description: 'Optional. Custom filename without extension (extension is auto-detected).',
            },
            outputDir: {
              type: 'string',
              description: 'Optional. Override the output directory (defaults to OS temp directory).',
            },
          },
          required: ['presentationId', 'pageObjectId', 'elementId'],
        },
      },
      {
        name: 'get_slide_notes',
        description: 'Get speaker notes for a specific slide. Returns just the notes text without other slide details.',
        inputSchema: {
          type: 'object',
          properties: {
            presentationId: {
              type: 'string',
              description: 'The ID of the presentation.',
            },
            pageObjectId: {
              type: 'string',
              description: 'The object ID of the slide to get notes from. NOTE: If copied from a URL like "slide=id.g123abc", remove the "id." prefix - use just "g123abc".',
            },
          },
          required: ['presentationId', 'pageObjectId'],
        },
      },
      {
        name: 'update_slide_notes',
        description:
          'Update speaker notes for a specific slide. PREFER append/insert/delete over replace to avoid losing existing content. Use get_slide_notes first to see what exists.',
        inputSchema: {
          type: 'object',
          properties: {
            presentationId: {
              type: 'string',
              description: 'The ID of the presentation.',
            },
            pageObjectId: {
              type: 'string',
              description: 'The object ID of the slide to update notes for. NOTE: If copied from a URL like "slide=id.g123abc", remove the "id." prefix - use just "g123abc".',
            },
            notes: {
              type: 'string',
              description: 'The text to insert/append/replace. Not used for delete operation.',
            },
            operation: {
              type: 'string',
              enum: ['replace', 'append', 'insert', 'delete'],
              description:
                'Operation: "append" (add to end - USE THIS for adding notes), "insert" (at position), "delete" (remove range), "replace" (ONLY use when rewriting ALL notes).',
            },
            insertionIndex: {
              type: 'number',
              description: 'Required for "insert" operation. Character position to insert text at (0-based).',
            },
            deleteStartIndex: {
              type: 'number',
              description: 'Required for "delete" operation. Start position of range to delete (0-based, inclusive).',
            },
            deleteEndIndex: {
              type: 'number',
              description: 'Required for "delete" operation. End position of range to delete (0-based, exclusive).',
            },
          },
          required: ['presentationId', 'pageObjectId', 'notes'],
        },
      },
      {
        name: 'get_slide_preview',
        description:
          'Get a visual preview (thumbnail) of the entire slide rendered as an image. Downloads thumbnail locally, then use Read(filePath) to visually analyze the slide layout, design, colors, and positioning. Use this for visual review; use get_page for structured data.',
        inputSchema: {
          type: 'object',
          properties: {
            presentationId: {
              type: 'string',
              description: 'The ID of the presentation.',
            },
            pageObjectId: {
              type: 'string',
              description: 'The object ID of the slide to preview. NOTE: If copied from a URL like "slide=id.g123abc", remove the "id." prefix - use just "g123abc".',
            },
            thumbnailSize: {
              type: 'string',
              enum: ['LARGE', 'MEDIUM', 'SMALL'],
              description: 'Optional. Thumbnail size (default: LARGE).',
            },
            mimeType: {
              type: 'string',
              enum: ['PNG', 'JPEG'],
              description: 'Optional. Image format (default: PNG).',
            },
            filename: {
              type: 'string',
              description: 'Optional. Custom filename without extension.',
            },
            outputDir: {
              type: 'string',
              description: 'Optional. Override output directory.',
            },
          },
          required: ['presentationId', 'pageObjectId'],
        },
      },
    ],
  }));

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;

    switch (name) {
      case 'create_presentation':
        return executeTool(slides, name, args, CreatePresentationArgsSchema, createPresentationTool);
      case 'get_presentation':
        return executeTool(slides, name, args, GetPresentationArgsSchema, getPresentationTool);
      case 'batch_update_presentation':
        return executeTool(slides, name, args, BatchUpdatePresentationArgsSchema, batchUpdatePresentationTool);
      case 'get_page':
        return executeTool(slides, name, args, GetPageArgsSchema, getPageTool);
      case 'summarize_presentation':
        return executeTool(slides, name, args, SummarizePresentationArgsSchema, summarizePresentationTool);
      case 'move_presentation': {
        if (!auth) {
          return {
            content: [{ type: 'text', text: 'Authentication object not available for move operation' }],
            isError: true,
          };
        }
        const moveTool = new MovePresentationTool(slides, auth);
        try {
          const validatedArgs = movePresentationSchema.parse(args);
          const result = await moveTool.execute(validatedArgs);
          return {
            content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
          };
        } catch (error) {
          return {
            content: [{ type: 'text', text: `Error: ${error instanceof Error ? error.message : String(error)}` }],
            isError: true,
          };
        }
      }
      case 'list_slide_images':
        return executeTool(slides, name, args, ListSlideImagesArgsSchema, listSlideImagesTool);
      case 'download_image':
        return executeTool(slides, name, args, DownloadImageArgsSchema, downloadImageTool);
      case 'get_slide_notes':
        return executeTool(slides, name, args, GetSlideNotesArgsSchema, getSlideNotesTool);
      case 'update_slide_notes':
        return executeTool(slides, name, args, UpdateSlideNotesArgsSchema, updateSlideNotesTool);
      case 'get_slide_preview':
        return executeTool(slides, name, args, GetSlidePreviewArgsSchema, getSlidePreviewTool);
      default:
        return {
          content: [{ type: 'text', text: `Unknown tool requested: ${name}` }],
          isError: true,
          errorCode: ErrorCode.MethodNotFound,
        };
    }
  });
};
