/**
 * 엑셀 열 문자를 숫자로 변환 (A=1, B=2, ..., Z=26, AA=27, etc.)
 */
export const columnLetterToNumber = (letter: string): number => {
  let col = 0;
  for (let i = 0; i < letter.length; i++) {
    col = col * 26 + (letter.charCodeAt(i) - 64);
  }
  return col;
};

/**
 * 엑셀 열 숫자를 문자로 변환 (1=A, 2=B, ..., 26=Z, 27=AA, etc.)
 */
export const columnNumberToLetter = (num: number): string => {
  let letter = '';
  while (num > 0) {
    const remainder = (num - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    num = Math.floor((num - 1) / 26);
  }
  return letter;
};

/**
 * 셀 주소를 파싱 (예: "A1" -> {row: 1, col: 1})
 */
export const parseCellAddress = (address: string): { row: number; col: number } | null => {
  const match = address.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return null;
  
  const col = columnLetterToNumber(match[1].toUpperCase());
  const row = parseInt(match[2], 10);
  
  return { row, col };
};

/**
 * 서명 placeholder로 사용되는 값들
 */
export const SIGNATURE_PLACEHOLDERS = ['1', '(1)', '1.', '1)', 'o', 'o)', '○'];

/**
 * 값이 서명 placeholder인지 확인
 */
export const isSignaturePlaceholder = (value: string): boolean => {
  return SIGNATURE_PLACEHOLDERS.includes(value.trim());
};
