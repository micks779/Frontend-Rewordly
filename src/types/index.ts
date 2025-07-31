export interface AnalysisResult {
  summary: string;
  actionItems: string[];
  priority: 'High' | 'Medium' | 'Low';
  sentiment: 'Positive' | 'Neutral' | 'Negative';
  category: string;
  rawAnalysis: string;
}

export interface RewordResult {
  rewordedText: string;
}

export interface ComposeResult {
  composedEmail: string;
} 