import React, { useState } from 'react';
import { 
  PrimaryButton,
  MessageBar, 
  MessageBarType, 
  Stack,
  IStackTokens,
  Text,
  mergeStyles
} from '@fluentui/react';
import TaskpaneLayout from './layout/TaskpaneLayout';
import { AnalysisResult } from '../types';

interface EmailAnalysisProps {
  onAnalysisComplete?: (result: AnalysisResult) => void;
  onGenerateResponse?: (result: AnalysisResult, originalEmail: string) => void;
}

// Styles
const analyzeButtonStyles = mergeStyles({
  width: '100%',
  marginBottom: '16px'
});

const generateResponseButtonStyles = mergeStyles({
  width: '100%',
  marginTop: '16px'
});

const sectionContainerStyles = mergeStyles({
  backgroundColor: '#f9f9fa',
  padding: '10px',
  margin: '0'
});

const sectionStyles = (sectionName: string) => {
  const sectionType = sectionName.toLowerCase();
  return mergeStyles({
    padding: '8px',
    marginBottom: '8px',
    borderRadius: '4px',
    backgroundColor: 
      sectionType.includes('context') ? '#f8f9fa' :
      sectionType.includes('key points') ? '#e7f3ff' :
      sectionType.includes('figures') ? '#fff8e1' :
      sectionType.includes('actions') ? '#e8f5e9' :
      sectionType.includes('attachments') ? '#f5e6ff' : '#ffffff',
    '&:last-child': {
      marginBottom: 0
    }
  });
};

const sectionTitleStyles = mergeStyles({
  fontSize: '11px',
  fontWeight: '600',
  color: '#323130',
  marginBottom: '4px'
});

const sectionContentStyles = mergeStyles({
  fontSize: '11px',
  lineHeight: '1.3',
  color: '#323130',
  whiteSpace: 'pre-wrap',
  '& p': {
    margin: '0 0 4px 0',
    '&:last-child': {
      marginBottom: 0
    }
  },
  '& ul': {
    margin: '2px 0',
    paddingLeft: '16px',
    listStyleType: 'disc'
  },
  '& li': {
    marginBottom: '2px',
    '&:last-child': {
      marginBottom: 0
    }
  }
});

const stackTokens: IStackTokens = { 
  childrenGap: 16
};

const EmailAnalysis: React.FC<EmailAnalysisProps> = ({ 
  onAnalysisComplete,
  onGenerateResponse 
}) => {
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<AnalysisResult | null>(null);
  const [originalEmail, setOriginalEmail] = useState<string>('');

  const analyzeEmail = async (emailContent: string, emailId: string): Promise<AnalysisResult> => {
    const response = await fetch(`${process.env.REACT_APP_API_URL}/api/analyze-email`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ emailContent, emailId }),
    });

    if (!response.ok) {
      throw new Error('Failed to analyze email');
    }

    const data = await response.json();
    return data;
  };

  const handleAnalyzeClick = async () => {
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

      setOriginalEmail(bodyResult);
      const analysis = await analyzeEmail(bodyResult, item.itemId);
      setResult(analysis);

      if (onAnalysisComplete) {
        onAnalysisComplete(analysis);
      }
    } catch (err) {
      console.error('Analysis error:', err);
      setError(err instanceof Error ? err.message : 'Failed to analyze email');
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleGenerateResponse = () => {
    if (result && originalEmail && onGenerateResponse) {
      onGenerateResponse(result, originalEmail);
    }
  };

  const renderAnalysisSections = (rawAnalysis: string) => {
    const sections = rawAnalysis.split('**').filter(Boolean);
    
    return sections.map((section, index) => {
      const lines = section.trim().split('\n');
      const title = lines[0].trim();
      const content = lines.slice(1).join('\n').trim();

      const formattedContent = content.split('\n').map(line => {
        const trimmedLine = line.trim();
        if (trimmedLine.startsWith('•')) {
          return `<li>${trimmedLine.substring(1).trim()}</li>`;
        }
        return `<p>${trimmedLine}</p>`;
      }).join('');

      const finalContent = formattedContent.includes('<li>') 
        ? `<ul>${formattedContent}</ul>`
        : formattedContent;

      return (
        <div key={index} className={sectionStyles(title)}>
          <div className={sectionTitleStyles}>{title}</div>
          <div 
            className={sectionContentStyles}
            dangerouslySetInnerHTML={{ __html: finalContent }}
          />
        </div>
      );
    });
  };

  return (
    <TaskpaneLayout>
      <Stack tokens={stackTokens}>
        {!result && (
          <PrimaryButton
            className={analyzeButtonStyles}
            text={isAnalyzing ? 'Analyzing...' : 'Analyze Email'}
            onClick={handleAnalyzeClick}
            disabled={isAnalyzing}
          />
        )}

        {error && (
          <MessageBar messageBarType={MessageBarType.error}>
            {error}
          </MessageBar>
        )}

        {result && result.rawAnalysis && (
          <div className={sectionContainerStyles}>
            {renderAnalysisSections(result.rawAnalysis)}
          </div>
        )}

        {result && (
          <PrimaryButton
            className={generateResponseButtonStyles}
            text="Generate Response"
            onClick={handleGenerateResponse}
          />
        )}
      </Stack>
    </TaskpaneLayout>
  );
};

export default EmailAnalysis;
