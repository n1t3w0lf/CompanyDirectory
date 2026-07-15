import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { IconButton, DefaultButton } from '@fluentui/react/lib/Button';

interface IPageButtonProps {
  page: number;
  active: boolean;
  onSelect: (page: number) => void;
}

/**
 * Single page-number button. Owns a stable click handler so the parent's JSX
 * stays free of inline arrow functions (react/jsx-no-bind).
 */
const PageButton: React.FC<IPageButtonProps> = ({ page, active, onSelect }) => {
  const handleClick = React.useCallback((): void => {
    onSelect(page);
  }, [page, onSelect]);

  return (
    <DefaultButton
      text={String(page)}
      onClick={handleClick}
      primary={active}
      checked={active}
      ariaLabel={`Page ${page}`}
      aria-current={active ? 'page' : undefined}
      styles={{ root: { minWidth: 36, padding: '0 8px' } }}
    />
  );
};

export interface IPaginationProps {
  currentPage: number;
  totalItems: number;
  pageSize: number;
  onPageChange: (page: number) => void;
  /** Number of page buttons to show around the current page. Default 5. */
  maxPageButtons?: number;
}

/**
 * Presentational client-side pager: previous / next, a windowed list of page
 * numbers, and an "X–Y of Z" summary. Renders nothing when there is a single page.
 */
export const Pagination: React.FC<IPaginationProps> = ({
  currentPage,
  totalItems,
  pageSize,
  onPageChange,
  maxPageButtons = 5
}) => {
  const totalPages = Math.max(1, Math.ceil(totalItems / pageSize));

  const goToPage = React.useCallback((page: number): void => {
    if (page >= 1 && page <= totalPages && page !== currentPage) {
      onPageChange(page);
    }
  }, [currentPage, totalPages, onPageChange]);

  const handlePrevious = React.useCallback((): void => {
    goToPage(currentPage - 1);
  }, [goToPage, currentPage]);

  const handleNext = React.useCallback((): void => {
    goToPage(currentPage + 1);
  }, [goToPage, currentPage]);

  // Nothing to page through
  if (totalPages <= 1) {
    return null;
  }

  // Build a windowed list of page numbers centred on the current page
  const half = Math.floor(maxPageButtons / 2);
  const endPage = Math.min(totalPages, Math.max(1, currentPage - half) + maxPageButtons - 1);
  const startPage = Math.max(1, endPage - maxPageButtons + 1);

  const pages: number[] = [];
  for (let p = startPage; p <= endPage; p++) {
    pages.push(p);
  }

  return (
    <Stack
      horizontal
      horizontalAlign="center"
      verticalAlign="center"
      wrap
      tokens={{ childrenGap: 12 }}
      role="navigation"
      aria-label="Pagination"
      styles={{ root: { marginTop: 16 } }}
    >
      <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
        <IconButton
          iconProps={{ iconName: 'ChevronLeft' }}
          title="Previous page"
          ariaLabel="Previous page"
          onClick={handlePrevious}
          disabled={currentPage === 1}
          allowDisabledFocus
        />

        {startPage > 1 && (
          <Text variant="small" style={{ padding: '0 4px', color: '#605E5C' }}>
            &hellip;
          </Text>
        )}

        {pages.map(page => (
          <PageButton
            key={page}
            page={page}
            active={page === currentPage}
            onSelect={goToPage}
          />
        ))}

        {endPage < totalPages && (
          <Text variant="small" style={{ padding: '0 4px', color: '#605E5C' }}>
            &hellip;
          </Text>
        )}

        <IconButton
          iconProps={{ iconName: 'ChevronRight' }}
          title="Next page"
          ariaLabel="Next page"
          onClick={handleNext}
          disabled={currentPage === totalPages}
          allowDisabledFocus
        />
      </Stack>
    </Stack>
  );
};
