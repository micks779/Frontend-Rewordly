import React, { useState } from 'react';
import { Stack, Pivot, PivotItem, IPivotStyles, IStackStyles } from '@fluentui/react';
import RewordText from './components/RewordText';
import ComposeEmail from './components/ComposeEmail';
import EmailAnalysis from './components/EmailAnalysis';
import { AnalysisResult } from './types';

const App: React.FC = () => {
  const [selectedTab, setSelectedTab] = useState<string>('reword');
  const [analysisContext, setAnalysisContext] = useState<AnalysisResult | null>(null);

  const handleAnalysisComplete = (result: AnalysisResult) => {
    setAnalysisContext(result);
  };

  // Root stack styles
  const stackStyles: IStackStyles = {
    root: {
      width: '100%',
      height: '100vh',
      overflow: 'hidden'
    }
  };

  // Custom styles for the Pivot component
  const pivotStyles: Partial<IPivotStyles> = {
    root: {
      backgroundColor: '#2b579a',
      width: '100%',
      '& .ms-Pivot': {
        display: 'flex',
        justifyContent: 'space-around',
        padding: '0',
        backgroundColor: '#2b579a'
      },
      '& button': {
        width: selectedTab === 'analyze' ? '33%' : '50%',
        maxWidth: 'none',
        minWidth: 'auto'
      }
    },
    link: {
      color: '#fff',
      padding: '6px 0',
      fontSize: '11px',
      selectors: {
        ':hover': {
          backgroundColor: '#366cc2'
        }
      }
    },
    linkIsSelected: {
      color: '#fff',
      padding: '6px 0',
      fontSize: '11px',
      backgroundColor: '#366cc2',
      selectors: {
        ':before': {
          backgroundColor: '#fff'
        }
      }
    },
    text: {
      fontSize: '11px',
      textAlign: 'center'
    }
  };

  return (
    <Stack styles={stackStyles}>
      <Stack.Item>
        <Pivot 
          selectedKey={selectedTab}
          onLinkClick={(item) => item && setSelectedTab(item.props.itemKey || 'reword')}
          styles={pivotStyles}
        >
          <PivotItem
            headerText="Reword"
            itemKey="reword"
          />
          <PivotItem
            headerText="Compose"
            itemKey="compose"
          />
          <PivotItem
            headerText="Analyze"
            itemKey="analyze"
          />
        </Pivot>
      </Stack.Item>
      <Stack.Item grow styles={{ root: { overflow: 'auto', height: 0 } }}>
        {selectedTab === 'reword' && (
          <RewordText />
        )}
        {selectedTab === 'compose' && (
          <ComposeEmail analysisContext={analysisContext} />
        )}
        {selectedTab === 'analyze' && (
          <EmailAnalysis onAnalysisComplete={handleAnalysisComplete} />
        )}
      </Stack.Item>
    </Stack>
  );
};

export default App; 