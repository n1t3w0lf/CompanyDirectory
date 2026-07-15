import * as React from 'react';
import { useCallback } from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import styles from './LetterIndex.module.scss';

export interface ILetterIndexProps {
  selectedLetter: string | null;
  onLetterSelect: (letter: string | null) => void;
  availableLetters?: Set<string>;
  /** Colour of the active/selected letter (from the property pane). */
  activeColor?: string;
  /** Colour of the normal/unselected letters (from the property pane). */
  normalColor?: string;
}

const ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');

export const LetterIndex: React.FC<ILetterIndexProps> = ({
  selectedLetter,
  onLetterSelect,
  availableLetters,
  activeColor = '#0078d4',
  normalColor = '#323130'
}) => {
  const handleLetterClick = useCallback((event: React.MouseEvent<HTMLButtonElement>): void => {
    const letter = event.currentTarget.dataset.letter;
    if (!letter) return;

    if (selectedLetter === letter) {
      // Clicking the same letter again clears the filter
      onLetterSelect(null);
    } else {
      onLetterSelect(letter);
    }
  }, [selectedLetter, onLetterSelect]);

  const handleAllClick = useCallback((): void => {
    onLetterSelect(null);
  }, [onLetterSelect]);

  return (
    <div className={styles.letterIndex}>
      <Stack horizontal tokens={{ childrenGap: 4 }} wrap>
        <button
          className={`${styles.letterButton} ${selectedLetter === null ? styles.active : ''}`}
          onClick={handleAllClick}
          aria-label="Show all users"
          title="Show all"
        >
          <Text variant="small" style={{ color: selectedLetter === null ? activeColor : normalColor }}>All</Text>
        </button>
        {ALPHABET.map(letter => {
          const isAvailable = !availableLetters || availableLetters.has(letter);
          const isSelected = selectedLetter === letter;

          return (
            <button
              key={letter}
              className={`${styles.letterButton} ${isSelected ? styles.active : ''} ${!isAvailable ? styles.disabled : ''}`}
              onClick={handleLetterClick}
              data-letter={letter}
              disabled={!isAvailable}
              aria-label={`Filter by ${letter}`}
              title={`Show names starting with ${letter}`}
            >
              <Text variant="small" style={{ color: isSelected ? activeColor : normalColor }}>{letter}</Text>
            </button>
          );
        })}
      </Stack>
    </div>
  );
};
