import { slides_v1 } from 'googleapis';

const SIZE_THRESHOLD = 100; // Points - elements smaller than this in both dimensions may be decorative

/**
 * Checks if an element has text content
 */
export const hasTextContent = (element: slides_v1.Schema$PageElement): boolean => {
  if (element.shape?.text?.textElements) {
    const textContent = element.shape.text.textElements
      .map((te) => te.textRun?.content?.trim() || '')
      .join('')
      .trim();
    return textContent.length > 0;
  }
  if (element.table?.tableRows) {
    for (const row of element.table.tableRows) {
      if (row.tableCells) {
        for (const cell of row.tableCells) {
          if (cell.text?.textElements) {
            const cellText = cell.text.textElements
              .map((te) => te.textRun?.content?.trim() || '')
              .join('')
              .trim();
            if (cellText.length > 0) return true;
          }
        }
      }
    }
  }
  return false;
};

/**
 * Checks if an element is decorative (should be filtered)
 */
export const isDecorativeElement = (element: slides_v1.Schema$PageElement): boolean => {
  // Check size threshold
  const width = element.size?.width?.magnitude ?? 0;
  const height = element.size?.height?.magnitude ?? 0;
  const isSmall = width < SIZE_THRESHOLD && height < SIZE_THRESHOLD;

  // Check if it's a semi-transparent shape (often decorative)
  const shapeProps = element.shape?.shapeProperties?.shapeBackgroundFill?.solidFill;
  const alpha = shapeProps?.alpha ?? 1;
  const isSemiTransparent = shapeProps !== undefined && alpha < 0.8;

  // Filter if small OR (semi-transparent AND no text)
  return isSmall || (isSemiTransparent && !hasTextContent(element));
};

/**
 * Filters out small decorative elements from page elements.
 * Keeps groups but filters their children.
 */
export const filterPageElements = (
  elements: slides_v1.Schema$PageElement[] | undefined
): slides_v1.Schema$PageElement[] => {
  if (!elements) return [];

  return elements
    .map((element) => {
      // If it's a group, keep the group but filter its children
      if (element.elementGroup?.children) {
        const originalChildCount = element.elementGroup.children.length;
        const filteredChildren = filterPageElements(element.elementGroup.children);

        // Keep the group with filtered children (may be empty)
        return {
          ...element,
          elementGroup: {
            ...element.elementGroup,
            children: filteredChildren,
          },
          // Add metadata about filtering if elements were removed
          ...(filteredChildren.length < originalChildCount && {
            description: `${element.description || 'Group'} (${originalChildCount - filteredChildren.length} decorative elements filtered)`,
          }),
        };
      }

      // Keep elements that have text or images
      if (hasTextContent(element) || element.image) {
        return element;
      }

      // Filter out decorative elements
      if (isDecorativeElement(element)) {
        return null;
      }

      // Keep other elements (tables, videos, etc.)
      return element;
    })
    .filter((el): el is slides_v1.Schema$PageElement => el !== null);
};
