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
  Checkbox
} from '@fluentui/react';
import TaskpaneLayout from './layout/TaskpaneLayout';
import { AnalysisResult, ComposeResult } from '../types';
import { copyToClipboard, replaceInOutlook } from '../utils/outlook';



interface ComposeEmailProps {
  analysisContext: AnalysisResult | null;
}

const stackTokens: IStackTokens = { 
  childrenGap: 12 
};

const buttonStackTokens: IStackTokens = { 
  childrenGap: 8 
};

const ComposeEmail: React.FC<ComposeEmailProps> = ({ analysisContext }) => {
  const [context, setContext] = useState('');
  const [useAnalysisContext, setUseAnalysisContext] = useState(false);
  const [isComposing, setIsComposing] = useState(false);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [composedEmail, setComposedEmail] = useState<string>('');
  const [currentEmailAnalysis, setCurrentEmailAnalysis] = useState<AnalysisResult | null>(null);

  const composeEmail = async () => {
    if (!context.trim()) {
      setError('Please enter context for the email');
      return;
    }

    setIsComposing(true);
    setError(null);

    try {
      let fullContext = context;
      
      // Add analysis context if requested and available
      if (useAnalysisContext && (analysisContext || currentEmailAnalysis)) {
        const analysisToUse = currentEmailAnalysis || analysisContext;
        if (analysisToUse) {
          const analysisText = analysisToUse.rawAnalysis || analysisToUse.summary || 'Email analyzed but no summary available';
          fullContext = `Email Analysis Context:\n${analysisText}\n\nUser Request: ${context}\n\nPlease compose an email that addresses the user's request while considering the context and content of the original email. Make the response relevant and contextual to the analyzed email.`;
        }
      }

      const response = await fetch(`${process.env.REACT_APP_API_URL}/api/compose`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          context: fullContext
        }),
      });

      if (!response.ok) {
        throw new Error('Failed to compose email');
      }

      const result: ComposeResult = await response.json();
      setComposedEmail(result.composedEmail);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to compose email');
    } finally {
      setIsComposing(false);
    }
  };

  const handleCopyToClipboard = async () => {
    try {
      await copyToClipboard(composedEmail);
      setError(null); // Clear any previous errors
      // You could add a success state here if needed
    } catch (err) {
      setError('Failed to copy to clipboard. Please try again.');
    }
  };

  const handleReplaceInOutlook = async () => {
    try {
      await replaceInOutlook(composedEmail);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to replace email in Outlook');
    }
  };



  const analyzeCurrentEmail = async () => {
    setIsAnalyzing(true);
    setError(null);

    try {
      const item = Office.context.mailbox.item;
      if (!item) {
        throw new Error('No email selected');
      }

      const bodyResult = await new Promise<string>((resolve, reject) => {
        item.body.getAsync(Office.CoercionType.Text, (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            reject(new Error(result.error.message));
          } else {
            resolve(result.value);
          }
        });
      });

      // Analyze the email using the same function as EmailAnalysis
      const response = await fetch(`${process.env.REACT_APP_API_URL}/api/analyze-email`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          emailContent: bodyResult,
          emailId: item.itemId
        }),
      });

      if (!response.ok) {
        throw new Error('Failed to analyze email');
      }

      const analysisResult: AnalysisResult = await response.json();
      console.log('Analysis result:', analysisResult); // Debug log
      setCurrentEmailAnalysis(analysisResult);
      setUseAnalysisContext(true);
    } catch (err) {
      console.error('Analysis error:', err);
      setError(err instanceof Error ? err.message : 'Failed to analyze current email');
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleContextChange = (value: string) => {
    setContext(value);
    // Auto-enable analysis context if user mentions "this email" or similar
    if (value.toLowerCase().includes('this email') || value.toLowerCase().includes('reply')) {
      setUseAnalysisContext(true);
    }
  };





  return (
    <TaskpaneLayout>
      <Stack tokens={stackTokens}>
        {error && (
          <MessageBar messageBarType={MessageBarType.error}>
            {error}
          </MessageBar>
        )}



        <Stack horizontal tokens={buttonStackTokens} horizontalAlign="space-between">
          <DefaultButton
            text={isAnalyzing ? 'Analyzing...' : 'Analyze Current Email'}
            onClick={analyzeCurrentEmail}
            disabled={isAnalyzing}
            iconProps={{ iconName: 'Search' }}
          />
        </Stack>

        <TextField
          label="What would you like to write?"
          multiline
          rows={4}
          value={context}
          onChange={(_, newValue) => handleContextChange(newValue || '')}
          placeholder="e.g., Write a follow-up email about the meeting yesterday, or Write a polite response to a complaint"
          styles={{
            field: {
              fontFamily: 'inherit',
            }
          }}
        />
        
        <div style={{ 
          fontSize: '10px', 
          backgroundColor: '#f0f6ff', 
          padding: '8px', 
          borderRadius: '4px',
          border: '1px solid #0078d4',
          color: '#605e5c'
        }}>
          💡 <strong>Tip:</strong> Type your request above, or use "Analyze Current Email" to get context for better composition.
        </div>

        {(analysisContext || currentEmailAnalysis) && (
          <Checkbox
            label="Include email analysis in composition"
            checked={useAnalysisContext}
            onChange={(_, checked) => setUseAnalysisContext(checked || false)}
            styles={{
              root: {
                fontSize: '11px'
              }
            }}
          />
        )}

        {(analysisContext || currentEmailAnalysis) && (
          <div style={{ 
            fontSize: '10px', 
            backgroundColor: '#f3f2f1', 
            padding: '8px', 
            borderRadius: '4px',
            color: '#605e5c'
          }}>
            <strong>Email Analysis Context:</strong><br/>
            {(currentEmailAnalysis || analysisContext)?.rawAnalysis || (currentEmailAnalysis || analysisContext)?.summary || 'Analysis completed but no context available.'}
          </div>
        )}

        <PrimaryButton
          text={isComposing ? 'Composing...' : 'Compose Email'}
          onClick={composeEmail}
          disabled={isComposing || (!context.trim())}
        />

        {isComposing && (
          <Stack horizontalAlign="center">
            <Spinner size={SpinnerSize.large} label="Composing your email..." />
          </Stack>
        )}

        {composedEmail && (
          <Stack tokens={stackTokens}>
            <TextField
              label="Composed Email"
              multiline
              rows={8}
              value={composedEmail}
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

export default ComposeEmail; 