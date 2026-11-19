import JSZip from 'jszip';

const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DRAWING_MAIN_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const SPREADSHEET_DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const PKG_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

/**
 * ISO/IEC 29500-1 §18.3.1.99 CT_Worksheet 스키마에서 <drawing> 요소 뒤에 와야 하는 요소들.
 * <drawing>이 이 요소들보다 뒤에 배치되면 Microsoft Excel이 즉시 복구 팝업을 띄운다.
 */
const POST_DRAWING_WORKSHEET_TAGS =
  /(<(?:legacyDrawing|legacyDrawingHF|drawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b)/;

/**
 * <worksheet> 루트 태그에 xmlns:r 네임스페이스 선언이 없으면 안전하게 추가
 */
export const ensureWorksheetRelationshipsNamespace = (sheetXml: string): string => {
  const wsOpenMatch = sheetXml.match(/<worksheet\b[^>]*>/);
  if (!wsOpenMatch) return sheetXml;
  const openTag = wsOpenMatch[0];
  if (/\bxmlns:r\s*=/.test(openTag)) return sheetXml;
  const updatedOpenTag = openTag.replace(/>$/, ` xmlns:r="${REL_NS}">`);
  return sheetXml.replace(openTag, updatedOpenTag);
};

/**
 * 워크시트 XML 내부에 <drawing r:id="..."/> 요소를 OOXML XSD 스키마 표준 위치에 삽입
 * - <legacyDrawing>, <tableParts>, <extLst> 등 후속 요소가 있으면 반드시 그 앞에 배치
 * - xmlns:r 선언이 누락된 워크시트면 자동 보강
 */
export const insertDrawingElementInWorksheetXml = (sheetXml: string, drawingRelId: string): string => {
  let xml = ensureWorksheetRelationshipsNamespace(sheetXml);
  if (/<drawing\b/.test(xml)) {
    return fixWorksheetDrawingOrder(xml);
  }

  const drawingTag = `<drawing r:id="${drawingRelId}"/>`;
  if (POST_DRAWING_WORKSHEET_TAGS.test(xml)) {
    return xml.replace(POST_DRAWING_WORKSHEET_TAGS, `${drawingTag}$1`);
  }

  return xml.replace('</worksheet>', `${drawingTag}</worksheet>`);
};

/**
 * 이미 존재하는 <drawing .../> 태그가 <extLst> 등 후속 요소 뒤에 잘못 배치되어 있다면
 * 표준 스키마 위치(후속 요소 앞)로 자동 재배치
 */
export const fixWorksheetDrawingOrder = (sheetXml: string): string => {
  const drawingMatch = sheetXml.match(/<drawing\b[^>]*\/>|<drawing\b[^>]*>[\s\S]*?<\/drawing>/);
  if (!drawingMatch) return sheetXml;

  const postMatch = sheetXml.match(POST_DRAWING_WORKSHEET_TAGS);
  if (!postMatch || postMatch.index === undefined) return sheetXml;

  const drawingIndex = drawingMatch.index ?? -1;
  if (drawingIndex === -1 || drawingIndex < postMatch.index) {
    return sheetXml; // 이미 올바른 순서
  }

  // 잘못된 위치의 <drawing>을 제거하고 첫 번째 후속 요소 바로 앞으로 이동
  const drawingTag = drawingMatch[0];
  const withoutDrawing = sheetXml.slice(0, drawingIndex) + sheetXml.slice(drawingIndex + drawingTag.length);
  return withoutDrawing.replace(POST_DRAWING_WORKSHEET_TAGS, `${drawingTag}$1`);
};

/**
 * .rels XML에 새로운 Relationship 태그들을 안전하게 추가
 * - null, 빈 파일, 자체 닫힘(<Relationships .../>), 일반 태그 모두 안전 처리
 */
export const appendRelationshipsXml = (relsXml: string | null, relationshipsToAdd: string): string => {
  if (!relationshipsToAdd) {
    return relsXml || `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PKG_RELS_NS}"></Relationships>`;
  }

  if (!relsXml || !relsXml.trim()) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PKG_RELS_NS}">${relationshipsToAdd}</Relationships>`;
  }

  // 자체 닫힘 <Relationships .../> 확장
  if (/<Relationships\b[^>]*\/>/.test(relsXml)) {
    return relsXml.replace(/<Relationships\b([^>]*)\/>/, `<Relationships$1>${relationshipsToAdd}</Relationships>`);
  }

  if (relsXml.includes('</Relationships>')) {
    return relsXml.replace('</Relationships>', `${relationshipsToAdd}</Relationships>`);
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PKG_RELS_NS}">${relationshipsToAdd}</Relationships>`;
};

/**
 * 기존 drawingN.xml의 루트 네임스페이스/접두사(xdr: 또는 기본 네임스페이스)를 감지하고
 * 자체 닫힘 확장 및 필수 네임스페이스(xmlns:a, xmlns:r) 선언을 보장
 */
export const prepareDrawingXml = (
  drawXml: string | null
): { xml: string; prefix: string; closeTag: string } => {
  if (!drawXml || !drawXml.trim()) {
    return {
      xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${SPREADSHEET_DRAWING_NS}" xmlns:a="${DRAWING_MAIN_NS}" xmlns:r="${REL_NS}"></xdr:wsDr>`,
      prefix: 'xdr:',
      closeTag: '</xdr:wsDr>',
    };
  }

  let xml = drawXml;

  // 루트 wsDr 태그 감지 (예: <xdr:wsDr ...> 또는 <wsDr ...>)
  const rootMatch = xml.match(/<([a-zA-Z0-9_]+:)?wsDr\b([^>]*?)(\/?)>/);
  if (!rootMatch) {
    return {
      xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${SPREADSHEET_DRAWING_NS}" xmlns:a="${DRAWING_MAIN_NS}" xmlns:r="${REL_NS}"></xdr:wsDr>`,
      prefix: 'xdr:',
      closeTag: '</xdr:wsDr>',
    };
  }

  const fullMatch = rootMatch[0];
  const prefix = rootMatch[1] || '';
  let attrs = rootMatch[2] || '';
  const isSelfClosing = rootMatch[3] === '/';
  const closeTag = `</${prefix}wsDr>`;

  if (!/\bxmlns:a\s*=/.test(attrs)) {
    attrs += ` xmlns:a="${DRAWING_MAIN_NS}"`;
  }
  if (!/\bxmlns:r\s*=/.test(attrs)) {
    attrs += ` xmlns:r="${REL_NS}"`;
  }

  if (isSelfClosing) {
    xml = xml.replace(fullMatch, `<${prefix}wsDr${attrs}>${closeTag}`);
  } else {
    xml = xml.replace(fullMatch, `<${prefix}wsDr${attrs}>`);
  }

  return { xml, prefix, closeTag };
};

/**
 * 워크북 탭 순서(xl/workbook.xml + xl/_rels/workbook.xml.rels)를 정확히 해석해
 * targetSheetIndex에 해당하는 실제 xl/worksheets/sheetN.xml 경로를 반환
 */
export const resolveTargetWorksheetPath = async (
  zip: JSZip,
  targetSheetIndex: number = 0
): Promise<string> => {
  try {
    const wbXmlFile = zip.file('xl/workbook.xml');
    const wbRelsFile = zip.file('xl/_rels/workbook.xml.rels');

    if (wbXmlFile && wbRelsFile) {
      const wbXml = await wbXmlFile.async('string');
      const wbRels = await wbRelsFile.async('string');

      // <sheet ... r:id="rIdX" .../> 순서대로 추출
      const sheetTags = [...wbXml.matchAll(/<sheet\b[^>]*>/g)].map(m => m[0]);
      const sheetRIds = sheetTags
        .map(tag => {
          const rIdMatch = tag.match(/\br:id="([^"]+)"/) || tag.match(/\bid="([^"]+)"/);
          return rIdMatch ? rIdMatch[1] : null;
        })
        .filter((id): id is string => !!id);

      const safeIdx = Math.max(0, Math.min(sheetRIds.length - 1, targetSheetIndex));
      const targetRId = sheetRIds[safeIdx];

      if (targetRId) {
        const relTags = [...wbRels.matchAll(/<Relationship\b[^>]*>/g)].map(m => m[0]);
        for (const relTag of relTags) {
          const idMatch = relTag.match(/\bId="([^"]+)"/);
          const targetMatch = relTag.match(/\bTarget="([^"]+)"/);
          if (idMatch && targetMatch && idMatch[1] === targetRId) {
            let target = targetMatch[1].replace(/\\/g, '/').replace(/^\/+/, '');
            if (target.startsWith('../')) {
              target = target.replace(/^\.\.\//, '');
            }
            const fullPath = target.startsWith('xl/') ? target : `xl/${target}`;
            if (zip.file(fullPath)) {
              return fullPath;
            }
          }
        }
      }
    }
  } catch (err) {
    console.warn('[resolveTargetWorksheetPath] workbook.xml 해석 경고, 숫자 정렬 폴백 사용:', err);
  }

  // 숫자 기반 자연 정렬 폴백 (sheet1, sheet2, ..., sheet10)
  const sheetFiles = Object.keys(zip.files)
    .filter(k => /^xl\/worksheets\/sheet\d+\.xml$/i.test(k))
    .sort((a, b) => {
      const numA = parseInt(a.match(/sheet(\d+)\.xml/i)?.[1] || '0', 10);
      const numB = parseInt(b.match(/sheet(\d+)\.xml/i)?.[1] || '0', 10);
      return numA - numB;
    });

  if (sheetFiles.length === 0) {
    return 'xl/worksheets/sheet1.xml';
  }
  const safeIdx = Math.max(0, Math.min(sheetFiles.length - 1, targetSheetIndex));
  return sheetFiles[safeIdx];
};

/**
 * ZIP 내에 존재하지 않는 고유한 새 drawingN.xml 경로 할당 (다중 시트 충돌 방지)
 */
export const allocateNextDrawingPath = (
  zip: JSZip
): { drawFilePath: string; drawRelsPath: string; drawFileName: string } => {
  const existingNums = Object.keys(zip.files)
    .map(k => k.match(/^xl\/drawings\/drawing(\d+)\.xml$/i))
    .filter((m): m is RegExpMatchArray => !!m)
    .map(m => parseInt(m[1], 10));

  const nextNum = Math.max(...existingNums, 0) + 1;
  const drawFileName = `drawing${nextNum}.xml`;
  return {
    drawFileName,
    drawFilePath: `xl/drawings/${drawFileName}`,
    drawRelsPath: `xl/drawings/_rels/${drawFileName}.rels`,
  };
};

/**
 * [Content_Types].xml에 png 확장자(대소문자 무관 중복 방지)와
 * 특정 drawingN.xml Override 등록을 보장
 */
export const ensureContentTypesForDrawingAndPng = (
  ctXml: string,
  drawFilePath: string
): string => {
  let ct = ctXml;

  // 1. PNG 확장자 대소문자 무관 중복 검사
  if (!/Extension="png"/i.test(ct)) {
    ct = ct.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>');
  }

  // 2. 해당 drawingN.xml PartName에 대한 Override 등록 보장
  const normalizedPartName = drawFilePath.startsWith('/') ? drawFilePath : `/${drawFilePath}`;
  const escapedPart = normalizedPartName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const overrideRegex = new RegExp(`PartName="${escapedPart}"`, 'i');

  if (!overrideRegex.test(ct)) {
    ct = ct.replace(
      '</Types>',
      `<Override PartName="${normalizedPartName}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`
    );
  }

  return ct;
};

/**
 * 생성된 XLSX 버퍼의 OOXML 구조적 무결성을 사전 검증 (다운로드 전 자가 진단)
 * - Microsoft Excel 복구 팝업을 유발하는 6대 치명 원인을 배포/다운로드 직전에 차단
 */
export const validateXlsxIntegrity = async (
  buffer: ArrayBuffer
): Promise<{ valid: boolean; issues: string[] }> => {
  const issues: string[] = [];

  try {
    const zip = await JSZip.loadAsync(buffer);

    const ctFile = zip.file('[Content_Types].xml');
    if (!ctFile) {
      issues.push('[Content_Types].xml 누락');
      return { valid: false, issues };
    }
    const ctXml = await ctFile.async('string');

    // 1. PNG 중복 Default 검사
    const pngDefaults = ctXml.match(/<Default\b[^>]*Extension="png"[^>]*>/gi) || [];
    if (pngDefaults.length > 1) {
      issues.push(`[Content_Types].xml에 PNG 확장자가 ${pngDefaults.length}회 중복 선언됨`);
    }

    // 2. 모든 워크시트의 <drawing> 순서 및 관계 무결성 검사
    for (const filename of Object.keys(zip.files)) {
      if (!/^xl\/worksheets\/sheet\d+\.xml$/i.test(filename)) continue;
      const sheetXml = await zip.file(filename)!.async('string');
      const drawingMatch = sheetXml.match(/<drawing\b[^>]*r:id="([^"]+)"[^>]*\/?>/);
      if (!drawingMatch) continue;

      const drawingIndex = drawingMatch.index ?? -1;
      const postMatch = sheetXml.match(POST_DRAWING_WORKSHEET_TAGS);
      if (postMatch && postMatch.index !== undefined && drawingIndex > postMatch.index) {
        issues.push(`${filename}: <drawing> 요소가 ${postMatch[1]} 뒤에 배치되어 OOXML 스키마 위반`);
      }

      const relId = drawingMatch[1];
      const baseName = filename.split('/').pop()!;
      const relsPath = `xl/worksheets/_rels/${baseName}.rels`;
      const relsFile = zip.file(relsPath);
      if (!relsFile) {
        issues.push(`${filename}: <drawing r:id="${relId}"/> 참조 대상 관계 파일(${relsPath}) 누락`);
        continue;
      }
      const relsXml = await relsFile.async('string');
      if (!new RegExp(`Id="${relId}"`).test(relsXml)) {
        issues.push(`${relsPath}: ${relId} 관계 항목 누락`);
      }
    }

    // 3. 모든 드로잉 XML의 스키마 및 이미지 관계 무결성 검사
    for (const filename of Object.keys(zip.files)) {
      if (!/^xl\/drawings\/drawing\d+\.xml$/i.test(filename)) continue;
      const drawXml = await zip.file(filename)!.async('string');

      // oneCellAnchor에 불법 editAs 속성 존재 여부 검사
      if (/<(?:[a-zA-Z0-9_]+:)?oneCellAnchor\b[^>]*\beditAs\s*=/.test(drawXml)) {
        issues.push(`${filename}: oneCellAnchor에 불법 editAs 속성 존재`);
      }

      // [Content_Types].xml 등록 여부 확인
      const escapedPart = (`/${filename}`).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      if (!new RegExp(`PartName="${escapedPart}"`, 'i').test(ctXml)) {
        issues.push(`[Content_Types].xml에 /${filename} Override 누락`);
      }

      // r:embed 이미지 참조 무결성 검사
      const embedIds = [...drawXml.matchAll(/r:embed="([^"]+)"/g)].map(m => m[1]);
      if (embedIds.length > 0) {
        const baseName = filename.split('/').pop()!;
        const drawRelsPath = `xl/drawings/_rels/${baseName}.rels`;
        const drawRelsFile = zip.file(drawRelsPath);
        if (!drawRelsFile) {
          issues.push(`${filename}: 이미지 관계 파일(${drawRelsPath}) 누락`);
          continue;
        }
        const drawRelsXml = await drawRelsFile.async('string');
        for (const rId of embedIds) {
          const relMatch = drawRelsXml.match(new RegExp(`<Relationship\\b[^>]*Id="${rId}"[^>]*>`));
          if (!relMatch) {
            issues.push(`${drawRelsPath}: 이미지 참조 ${rId} 누락`);
            continue;
          }
          const targetMatch = relMatch[0].match(/Target="([^"]+)"/);
          if (targetMatch) {
            const mediaTarget = targetMatch[1].replace(/^\.\.\//, 'xl/').replace(/^\/+/, '');
            const fullMediaPath = mediaTarget.startsWith('xl/') ? mediaTarget : `xl/drawings/${mediaTarget}`;
            if (!zip.file(fullMediaPath)) {
              issues.push(`${filename}: 미디어 파일(${fullMediaPath}) 누락`);
            }
          }
        }
      }
    }

    return { valid: issues.length === 0, issues };
  } catch (err) {
    issues.push(`ZIP/XML 무결성 검사 중 예외 발생: ${err instanceof Error ? err.message : String(err)}`);
    return { valid: false, issues };
  }
};

