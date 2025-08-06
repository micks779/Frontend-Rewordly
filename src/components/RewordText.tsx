import React, { useState } from 'react';
import { 
  Stack, 
  TextField, 
  PrimaryButton, 
  DefaultButton,
  MessageBar, 
  MessageBarType,
  Spinner,
  SpinnerSize,
  IStackTokens,
  mergeStyles
} from '@fluentui/react';
import TaskpaneLayout from './layout/TaskpaneLayout';
import { RewordResult } from '../types';
import { copyToClipboard, replaceInOutlook } from '../utils/outlook';

const tonePresets = [
  { key: 'professional', label: 'Professional' },
  { key: 'warm', label: 'Warm' },
  { key: 'concise', label: 'Concise' },
  { key: 'formal', label: 'Formal' },
  { key: 'casual', label: 'Casual' },
  { key: 'friendly', label: 'Friendly' }
];

const stackTokens: IStackTokens = { 
  childrenGap: 12 
};

const buttonStackTokens: IStackTokens = { 
  childrenGap: 8 
};

const toneButtonStyles = mergeStyles({
  minWidth: '80px',
  height: '28px',
  fontSize: '10px'
});

const selectedToneButtonStyles = mergeStyles({
  minWidth: '80px',
  height: '28px',
  fontSize: '10px',
  backgroundColor: '#2b579a',
  color: 'white',
  selectors: {
    ':hover': {
      backgroundColor: '#366cc2'
    }
  }
});

const RewordText: React.FC = () => {
  const [originalText, setOriginalText] = useState('');
  const [selectedTone, setSelectedTone] = useState<string>('');
  const [customInstructions, setCustomInstructions] = useState('');
  const [isRewording, setIsRewording] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [rewordedText, setRewordedText] = useState<string>('');

  const rewordText = async () => {
    if (!originalText.trim()) {
      setError('Please enter text to reword');
      return;
    }

    if (!selectedTone && !customInstructions.trim()) {
      setError('Please select a tone or enter custom instructions');
      return;
    }

    setIsRewording(true);
    setError(null);

    try {
      // Create request body based on user selection
      const requestBody = selectedTone
        ? { selectedText: originalText, toneInstructions: `Make this text more ${selectedTone}` }
        : { selectedText: originalText, toneInstructions: customInstructions.trim() };
      
      const response = await fetch(`${process.env.REACT_APP_API_URL}/api/reword`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        throw new Error('Failed to reword text');
      }

      const result: RewordResult = await response.json();
      setRewordedText(result.rewording_text);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to reword text');
    } finally {
      setIsRewording(false);
    }
  };

  const handleCopyToClipboard = async () => {
    try {
      await copyToClipboard(rewordedText);
      // Show success message
      setError(null); // Clear any previous errors
      // You could add a success state here if needed
    } catch (err) {
      setError('Failed to copy to clipboard. Please try again.');
    }
  };

  const handleReplaceInOutlook = async () => {
    try {
      await replaceInOutlook(rewordedText);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to replace text in Outlook');
    }
  };

  const handleToneSelect = (tone: string) => {
    setSelectedTone(tone);
    setCustomInstructions(''); // Clear custom instructions when tone is selected
  };

  const handleCustomInstructionsChange = (value: string) => {
    setCustomInstructions(value);
    setSelectedTone(''); // Clear selected tone when custom instructions are entered
  };

  return (
    <TaskpaneLayout>
      <Stack tokens={stackTokens}>
        {error && (
          <MessageBar messageBarType={MessageBarType.error}>
            {error}
          </MessageBar>
        )}

        <TextField
          label="Text to Reword"
          multiline
          rows={6}
          value={originalText}
          onChange={(_, newValue) => setOriginalText(newValue || '')}
          placeholder="Paste your email text here..."
          styles={{
            field: {
              fontFamily: 'inherit',
            }
          }}
        />

        <Stack>
          <div style={{ fontSize: '11px', fontWeight: '600', marginBottom: '8px' }}>
            Select Tone:
          </div>
          <Stack horizontal tokens={buttonStackTokens} wrap>
            {tonePresets.map((tone) => (
              <DefaultButton
                key={tone.key}
                text={tone.label}
                onClick={() => handleToneSelect(tone.key)}
                className={selectedTone === tone.key ? selectedToneButtonStyles : toneButtonStyles}
              />
            ))}
          </Stack>
        </Stack>

        <TextField
          label="Or Custom Instructions"
          value={customInstructions}
          onChange={(_, newValue) => handleCustomInstructionsChange(newValue || '')}
          placeholder="e.g., Make this more polite and professional"
          styles={{
            field: {
              fontFamily: 'inherit',
            }
          }}
        />

        <PrimaryButton
          text={isRewording ? 'Rewording...' : 'Reword Text'}
          onClick={rewordText}
          disabled={isRewording || (!originalText.trim())}
        />

        {isRewording && (
          <Stack horizontalAlign="center">
            <Spinner size={SpinnerSize.large} label="Rewording your text..." />
          </Stack>
        )}

        {rewordedText && (
          <Stack tokens={stackTokens}>
            <TextField
              label="Reworded Text"
              multiline
              rows={6}
              value={rewordedText}
              readOnly
              styles={{
                field: {
                  fontFamily: 'inherit',
                }
              }}
            />
            
            <Stack horizontal tokens={buttonStackTokens}>
              <DefaultButton
                text="Copy"
                onClick={() => handleCopyToClipboard()}
                iconProps={{ iconName: 'Copy' }}
              />
              <PrimaryButton
                text="Replace in Outlook"
                onClick={() => handleReplaceInOutlook()}
                iconProps={{ iconName: 'Replace' }}
              />
            </Stack>
          </Stack>
        )}
      </Stack>
    </TaskpaneLayout>
  );
};

export default RewordText; 