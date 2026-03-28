export interface ResolvedColorToken {
  kind: 'rgb' | 'theme' | 'automatic';
  value: string;
  source?: string;
}

export interface ResolvedFontToken {
  family: string;
  fallbackFamilies: string[];
  source?: string;
}

export interface ResolvedTextStyle {
  color?: ResolvedColorToken;
  font?: ResolvedFontToken;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
}

export interface ResolvedTableCellStyle {
  backgroundColor?: ResolvedColorToken;
  borderColor?: ResolvedColorToken;
  textStyle?: ResolvedTextStyle;
}

export interface ResolvedDrawingStyle {
  fillColor?: ResolvedColorToken;
  strokeColor?: ResolvedColorToken;
  textStyle?: ResolvedTextStyle;
}

export interface ResolvedAnnotationStyle {
  accentColor?: ResolvedColorToken;
  authorLabel?: string;
}
