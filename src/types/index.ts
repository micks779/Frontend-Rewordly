export interface AnalysisResult {
  summary: string;
  actionItems: string[];
  priority: 'High' | 'Medium' | 'Low';
  sentiment: 'Positive' | 'Neutral' | 'Negative';
  category: string;
  rawAnalysis: string;
}

export interface RewordResult {
  success: boolean;
  rewording_text: string;
  original_text: string;
  tone_instructions: string;
}

export interface ComposeResult {
  success: boolean;
  composed_email: string;
  composition_instructions: string;
} 