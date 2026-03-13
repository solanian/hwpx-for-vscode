import { XMLParser } from 'fast-xml-parser';
import JSZip from 'jszip';

interface BorderInfo {
    type: string;   // NONE, SOLID, DOTTED, DASHED, DOUBLE 등
    width: string;  // "0.12 mm" 등
    color: string;  // "#000000" 등
}

interface BorderFillEntry {
    faceColor: string;
    fillCss: string;  // CSS background 속성 (gradient, image 등 포함)
    topBorder: BorderInfo;
    bottomBorder: BorderInfo;
    leftBorder: BorderInfo;
    rightBorder: BorderInfo;
}

export interface ParseResult {
    html: string;
    css: string;
    metadata: {
        title?: string;
        creator?: string;
        subject?: string;
        lastSavedBy?: string;
        createdDate?: string;
        modifiedDate?: string;
    }
}

export class HwpxParser {
    private static borderFillMap: Record<string, BorderFillEntry> = {};
    // fontMap: { lang -> { id -> faceName } }
    private static fontMap: Record<string, Record<string, string>> = {};
    // numberingMap: { id -> { level -> { numFormat, start, textPattern, autoIndent } } }
    private static numberingMap: Record<string, Record<string, { numFormat: string; start: number; textPattern: string; autoIndent: boolean }>> = {};
    // autoNumCounters: { numType -> currentCount }
    private static autoNumCounters: Record<string, number> = {};
    // styleMap: { id -> { paraPrIDRef, charPrIDRef, type } }
    private static styleMap: Record<string, { paraPrIDRef: string; charPrIDRef: string; type: string }> = {};
    // tabPrMap: { id -> TabItem[] }
    private static tabPrMap: Record<string, { pos: number; type: string; leader: string }[]> = {};

    static async parse(zip: any): Promise<ParseResult> {
        const result: ParseResult = {
            html: '',
            css: '',
            metadata: {}
        };
        this.borderFillMap = {};
        this.fontMap = {};
        this.numberingMap = {};
        this.autoNumCounters = {};
        this.styleMap = {};
        this.tabPrMap = {};

        // 이미지 바이너리 맵 (id -> data URI)
        const imageMap: Record<string, string> = {};

        const parser = new XMLParser({
            ignoreAttributes: false,
            attributeNamePrefix: "@_",
            textNodeName: "#text",
            trimValues: false,
            isArray: (name: string, jpath: string, isLeafNode: boolean, isAttribute: boolean) => {
                if([
                    'hp:p', 'hp:run', 'hp:t', 'opf:meta', 'hh:charPr', 'hh:paraPr',
                    'hp:tr', 'hp:tc', 'opf:item',
                    'hh:fontface', 'hh:font', 'hh:borderFill', 'hh:paraHead',
                    'hp:pageBorderFill', 'hp:stringParam', 'hh:numbering', 'hh:tabItem',
                    'hh:style', 'hp:colSz', 'hc:color', 'hp:pt', 'hp:seg', 'hp:listItem'
                ].includes(name)) return true;
                return false;
            }
        });

        // 1. 문서 메타데이터 및 이미지 바이너리 맵 파싱
        try {
            const hpfXml = await zip.file('Contents/content.hpf')?.async('string');
            if (hpfXml) {
                const hpfObj = parser.parse(hpfXml);

                // 메타데이터 처리
                const metadata = hpfObj['opf:package']?.['opf:metadata'];
                if (metadata) {
                    result.metadata.title = metadata['opf:title'];
                    const metas = metadata['opf:meta'];
                    if (metas && Array.isArray(metas)) {
                        for (const meta of metas) {
                            const name = meta['@_name'];
                            const text = meta['#text'];
                            if (name === 'creator') result.metadata.creator = text;
                            else if (name === 'subject') result.metadata.subject = text;
                            else if (name === 'lastsaveby') result.metadata.lastSavedBy = text;
                            else if (name === 'CreatedDate') result.metadata.createdDate = text;
                            else if (name === 'ModifiedDate') result.metadata.modifiedDate = text;
                        }
                    }
                }

                // 빈데이터(BinData) 매핑 구성
                const manifest = hpfObj['opf:package']?.['opf:manifest'];
                if (manifest && manifest['opf:item']) {
                    const items = Array.isArray(manifest['opf:item']) ? manifest['opf:item'] : [manifest['opf:item']];
                    for (const item of items) {
                        const href = item['@_href'];
                        const id = item['@_id'];
                        const mediaType = item['@_media-type'];

                        if (href && href.startsWith('BinData/') && mediaType && mediaType.startsWith('image/')) {
                            const fileObj = zip.file(href);
                            if (fileObj) {
                                const base64 = await fileObj.async('base64');
                                imageMap[id] = `data:${mediaType};base64,${base64}`;
                            }
                        }
                    }
                }
            }
        } catch (e) {
            console.error('Error parsing content.hpf', e);
        }

        let defaultPageCss = `width: 210mm; min-height: 297mm; padding: 30mm 20mm; box-sizing: border-box; position: relative; overflow: hidden;`;

        // 2. 스타일(header.xml) 파싱
        try {
            const headerXml = await zip.file('Contents/header.xml')?.async('string');
            if (headerXml) {
                const headerObj = parser.parse(headerXml);

                // 기본 페이지 사이즈 (secDef > pageDef)
                const head = headerObj['hh:head'];
                if (head) {
                    const secDef = this.findDeepNode(head, 'hh:secDef') || this.findDeepNode(head, 'hp:secDef');
                    const pageDef = secDef && (secDef['hh:pageDef'] || secDef['hp:pageDef']);
                    if (pageDef) {
                        let widthMm = parseInt(pageDef['@_width'] || '59528') / 283.465;
                        let heightMm = parseInt(pageDef['@_height'] || '84188') / 283.465;
                        const landscapeVal = String(pageDef['@_landscape'] || '').trim().toLowerCase();
                        // landscape 속성이 있으면 (NARROWLY/WIDELY 등) 가로 방향
                        const isLandscape = landscapeVal !== '' && landscapeVal !== '0' && landscapeVal !== 'false' && landscapeVal !== 'undefined';

                        if (isLandscape && widthMm < heightMm) {
                            const temp = widthMm;
                            widthMm = heightMm;
                            heightMm = temp;
                        }

                        const marginTag = pageDef['hh:pageMar'] || pageDef['hp:pageMar'] || pageDef['hh:margin'] || pageDef['hp:margin'];
                        let paddingCss = 'padding: 30mm 20mm;';
                        if (marginTag) {
                            // top+header, bottom+footer = 본문 영역까지의 전체 여백
                            const topMm = (parseInt(marginTag['@_top'] || '0') + parseInt(marginTag['@_header'] || '0')) / 283.465;
                            const bottomMm = (parseInt(marginTag['@_bottom'] || '0') + parseInt(marginTag['@_footer'] || '0')) / 283.465;
                            const leftMm = parseInt(marginTag['@_left'] || '0') / 283.465;
                            const rightMm = parseInt(marginTag['@_right'] || '0') / 283.465;
                            paddingCss = `padding: ${topMm}mm ${rightMm}mm ${bottomMm}mm ${leftMm}mm;`;
                        }

                        defaultPageCss = `width: ${widthMm}mm; min-height: ${heightMm}mm; ${paddingCss} box-sizing: border-box; position: relative; overflow: hidden;`;
                    }
                }

                const refList = head?.['hh:refList'];
                if (refList) {
                    // ===== fontMap 구축 =====
                    const fontfaces = refList['hh:fontfaces'];
                    if (fontfaces) {
                        const ffList = Array.isArray(fontfaces['hh:fontface']) ? fontfaces['hh:fontface'] : (fontfaces['hh:fontface'] ? [fontfaces['hh:fontface']] : []);
                        for (const ff of ffList) {
                            const lang = (ff['@_lang'] || '').toUpperCase();
                            const fonts = Array.isArray(ff['hh:font']) ? ff['hh:font'] : (ff['hh:font'] ? [ff['hh:font']] : []);
                            const langMap: Record<string, string> = {};
                            for (const font of fonts) {
                                langMap[String(font['@_id'])] = font['@_face'] || '';
                            }
                            this.fontMap[lang] = langMap;
                        }
                    }

                    // ===== charPr CSS 생성 =====
                    const charProps = refList['hh:charProperties']?.['hh:charPr'];
                    if (charProps) {
                        for (const cp of charProps) {
                            const id = cp['@_id'];
                            let css = '';

                            // 글꼴 크기
                            const heightVal = parseInt(cp['@_height'] || '0');
                            let fontSize = heightVal / 100;

                            // 상대 크기 보정 (relSz)
                            const relSz = cp['hh:relSz'];
                            if (relSz) {
                                const relVal = parseInt(relSz['@_hangul'] || '100');
                                if (relVal !== 100 && fontSize > 0) {
                                    fontSize = fontSize * relVal / 100;
                                }
                            }
                            if (fontSize > 0) css += `font-size: ${fontSize}pt; `;

                            // 글꼴 색상
                            if (cp['@_textColor'] && cp['@_textColor'] !== '#000000' && cp['@_textColor'] !== 'none') {
                                css += `color: ${cp['@_textColor']}; `;
                            }

                            // 굵게
                            if (cp['hh:bold'] !== undefined || cp['@_bold'] === '1' || cp['@_bold'] === 'true') css += `font-weight: bold; `;

                            // 기울임
                            if (cp['hh:italic'] !== undefined || cp['@_italic'] === '1' || cp['@_italic'] === 'true') css += `font-style: italic; `;

                            // 밑줄 + 취소선 (합쳐서 text-decoration)
                            let textDeco = '';
                            let textDecoColor = '';
                            if (cp['hh:underline'] && cp['hh:underline']['@_type'] !== 'NONE') {
                                textDeco += 'underline ';
                                if (cp['hh:underline']['@_color']) textDecoColor = cp['hh:underline']['@_color'];
                            }
                            if (cp['hh:strikeout'] && cp['hh:strikeout']['@_shape'] !== 'NONE') {
                                textDeco += 'line-through ';
                                if (!textDecoColor && cp['hh:strikeout']['@_color']) {
                                    textDecoColor = cp['hh:strikeout']['@_color'];
                                }
                            }
                            if (textDeco.trim()) {
                                css += `text-decoration: ${textDeco.trim()}; `;
                                if (textDecoColor && textDecoColor !== '#000000') css += `text-decoration-color: ${textDecoColor}; `;
                            }

                            // 글꼴 이름 (fontRef → fontMap)
                            const fontRef = cp['hh:fontRef'];
                            if (fontRef) {
                                const hangulFontId = String(fontRef['@_hangul'] || '');
                                const latinFontId = String(fontRef['@_latin'] || '');
                                const hangulFontName = this.fontMap['HANGUL']?.[hangulFontId] || '';
                                const latinFontName = this.fontMap['LATIN']?.[latinFontId] || '';
                                if (hangulFontName) {
                                    const fonts: string[] = [`"${hangulFontName}"`];
                                    if (latinFontName && latinFontName !== hangulFontName) {
                                        fonts.push(`"${latinFontName}"`);
                                    }
                                    fonts.push('sans-serif');
                                    css += `font-family: ${fonts.join(', ')}; `;
                                }
                            }

                            // 자간 (spacing)
                            const spacing = cp['hh:spacing'];
                            if (spacing) {
                                const spacingVal = parseInt(spacing['@_hangul'] || '0');
                                if (spacingVal !== 0) {
                                    css += `letter-spacing: ${spacingVal / 100}em; `;
                                }
                            }

                            // 장평 (ratio - scaleX)
                            const ratio = cp['hh:ratio'];
                            if (ratio) {
                                const ratioVal = parseInt(ratio['@_hangul'] || '100');
                                if (ratioVal !== 100) {
                                    css += `display: inline-block; transform: scaleX(${ratioVal / 100}); `;
                                }
                            }

                            // 그림자 (shadow)
                            const shadow = cp['hh:shadow'];
                            if (shadow && shadow['@_type'] !== 'NONE') {
                                const sColor = shadow['@_color'] || '#C0C0C0';
                                const sX = parseInt(shadow['@_offsetX'] || '10') / 10;
                                const sY = parseInt(shadow['@_offsetY'] || '10') / 10;
                                css += `text-shadow: ${sX}px ${sY}px ${sColor}; `;
                            }

                            // 외곽선 (outline)
                            const outline = cp['hh:outline'];
                            if (outline && outline['@_type'] !== 'NONE') {
                                css += `-webkit-text-stroke: 0.06em currentColor; `;
                            }

                            // 수직 오프셋 (offset)
                            const offset = cp['hh:offset'];
                            if (offset) {
                                const offsetVal = parseInt(offset['@_hangul'] || '0');
                                if (offsetVal !== 0) {
                                    css += `vertical-align: ${offsetVal}%; `;
                                }
                            }

                            // 글자 배경색 (shadeColor)
                            const shadeColor = cp['@_shadeColor'];
                            if (shadeColor && shadeColor !== 'none') {
                                css += `background-color: ${shadeColor}; `;
                            }

                            result.css += `.char-${id} { ${css} }\n`;
                        }
                    }

                    // ===== 글머리 기호(bullet) 정의 =====
                    const bulletMap: Record<string, { char: string; autoIndent: boolean }> = {};
                    const bulletDefs = refList['hh:bullets'];
                    if (bulletDefs) {
                        const bullets = Array.isArray(bulletDefs['hh:bullet']) ? bulletDefs['hh:bullet'] : (bulletDefs['hh:bullet'] ? [bulletDefs['hh:bullet']] : []);
                        for (const b of bullets) {
                            const bid = b['@_id'];
                            let bchar = b['@_char'] || '•';
                            if (bchar && bchar.charCodeAt(0) >= 0xE000 && bchar.charCodeAt(0) <= 0xF8FF) {
                                bchar = '•';
                            }
                            bchar = bchar || '•';
                            // paraHead의 autoIndent 속성 추출
                            const paraHeads = Array.isArray(b['hh:paraHead']) ? b['hh:paraHead'] : (b['hh:paraHead'] ? [b['hh:paraHead']] : []);
                            const autoIndent = paraHeads.length > 0 && (paraHeads[0]['@_autoIndent'] === '1' || paraHeads[0]['@_autoIndent'] === 'true');
                            bulletMap[bid] = { char: bchar, autoIndent };
                        }
                    }

                    // ===== numberingMap 구축 =====
                    const numberingDefs = refList['hh:numberings'];
                    if (numberingDefs) {
                        const numberings = Array.isArray(numberingDefs['hh:numbering']) ? numberingDefs['hh:numbering'] : (numberingDefs['hh:numbering'] ? [numberingDefs['hh:numbering']] : []);
                        for (const n of numberings) {
                            const nid = n['@_id'];
                            const levelMap: Record<string, { numFormat: string; start: number; textPattern: string; autoIndent: boolean }> = {};
                            const paraHeads = Array.isArray(n['hh:paraHead']) ? n['hh:paraHead'] : (n['hh:paraHead'] ? [n['hh:paraHead']] : []);
                            for (const ph of paraHeads) {
                                const level = ph['@_level'] || '1';
                                levelMap[level] = {
                                    numFormat: ph['@_numFormat'] || 'DIGIT',
                                    start: parseInt(ph['@_start'] || '1'),
                                    textPattern: ph['#text'] || '',
                                    autoIndent: ph['@_autoIndent'] === '1' || ph['@_autoIndent'] === 'true',
                                };
                            }
                            this.numberingMap[nid] = levelMap;
                        }
                    }

                    // ===== tabPrMap 구축 =====
                    const tabProps = refList['hh:tabProperties'];
                    if (tabProps) {
                        const tabPrs = Array.isArray(tabProps['hh:tabPr']) ? tabProps['hh:tabPr'] : (tabProps['hh:tabPr'] ? [tabProps['hh:tabPr']] : []);
                        for (const tp of tabPrs) {
                            const tpId = tp['@_id'];
                            const items = Array.isArray(tp['hh:tabItem']) ? tp['hh:tabItem'] : (tp['hh:tabItem'] ? [tp['hh:tabItem']] : []);
                            this.tabPrMap[tpId] = items.map((ti: any) => ({
                                pos: parseInt(ti['@_pos'] || '0'),
                                type: ti['@_type'] || 'LEFT',
                                leader: ti['@_leader'] || 'NONE',
                            }));
                        }
                    }

                    // ===== borderFill 정의 (셀 테두리 + 배경색) =====
                    const borderFillProps = refList['hh:borderFills'];
                    if (borderFillProps) {
                        const bfs = Array.isArray(borderFillProps['hh:borderFill']) ? borderFillProps['hh:borderFill'] : (borderFillProps['hh:borderFill'] ? [borderFillProps['hh:borderFill']] : []);
                        for (const bf of bfs) {
                            const bfId = bf['@_id'];
                            const fillBrush = bf['hc:fillBrush'];
                            let faceColor = 'none';
                            let fillCss = '';

                            if (fillBrush) {
                                // 1) 단색 채움 (winBrush)
                                if (fillBrush['hc:winBrush']) {
                                    faceColor = fillBrush['hc:winBrush']['@_faceColor'] || 'none';
                                    if (faceColor && faceColor !== 'none') {
                                        fillCss = `background-color: ${faceColor};`;
                                    }
                                }
                                // 2) 그라데이션 채움 (gradation)
                                if (fillBrush['hc:gradation']) {
                                    const grad = fillBrush['hc:gradation'];
                                    const gradType = (grad['@_type'] || 'LINEAR').toUpperCase();
                                    const angle = parseInt(grad['@_angle'] || '90');
                                    const colors = Array.isArray(grad['hc:color']) ? grad['hc:color'] : (grad['hc:color'] ? [grad['hc:color']] : []);
                                    const colorStops = colors.map((c: any) => c['@_value'] || '#FFFFFF');

                                    if (colorStops.length >= 2) {
                                        if (gradType === 'RADIAL' || gradType === 'CONICAL') {
                                            const cx = parseInt(grad['@_centerX'] || '50');
                                            const cy = parseInt(grad['@_centerY'] || '50');
                                            fillCss = `background: radial-gradient(circle at ${cx}% ${cy}%, ${colorStops.join(', ')});`;
                                        } else {
                                            fillCss = `background: linear-gradient(${angle}deg, ${colorStops.join(', ')});`;
                                        }
                                    }
                                }
                                // 3) 이미지 채움 (imgBrush)
                                if (fillBrush['hc:imgBrush']) {
                                    const imgBrush = fillBrush['hc:imgBrush'];
                                    const mode = (imgBrush['@_mode'] || 'TILE').toUpperCase();
                                    const imgNode = imgBrush['hc:img'];
                                    if (imgNode) {
                                        const binRef = imgNode['@_binaryItemIDRef'];
                                        if (binRef && imageMap[binRef]) {
                                            if (mode === 'TILE') {
                                                fillCss = `background-image: url(${imageMap[binRef]}); background-repeat: repeat; background-size: auto;`;
                                            } else if (mode === 'CENTER') {
                                                fillCss = `background-image: url(${imageMap[binRef]}); background-repeat: no-repeat; background-position: center;`;
                                            } else {
                                                // TOTAL, FIT 등 → stretch
                                                fillCss = `background-image: url(${imageMap[binRef]}); background-size: 100% 100%; background-repeat: no-repeat;`;
                                            }
                                        }
                                    }
                                }
                            }

                            const parseBorder = (node: any): BorderInfo => ({
                                type: node?.['@_type'] || 'NONE',
                                width: node?.['@_width'] || '0.1 mm',
                                color: node?.['@_color'] || '#000000',
                            });

                            this.borderFillMap[bfId] = {
                                faceColor,
                                fillCss,
                                topBorder: parseBorder(bf['hh:topBorder']),
                                bottomBorder: parseBorder(bf['hh:bottomBorder']),
                                leftBorder: parseBorder(bf['hh:leftBorder']),
                                rightBorder: parseBorder(bf['hh:rightBorder']),
                            };
                        }
                    }

                    // ===== paraPr CSS 생성 =====
                    const paraProps = refList['hh:paraProperties']?.['hh:paraPr'];
                    if (paraProps) {
                        for (const pp of paraProps) {
                            const id = pp['@_id'];
                            let css = '';

                            // 분할 설정 (breakSetting)
                            const breakSetting = pp['hh:breakSetting'];
                            if (breakSetting) {
                                if (breakSetting['@_pageBreakBefore'] === '1') css += `break-before: page; `;
                                if (breakSetting['@_keepWithNext'] === '1') css += `break-after: avoid; `;
                                if (breakSetting['@_widowOrphan'] === '1') css += `orphans: 2; widows: 2; `;
                            }

                            // 자동 간격 (autoSpacing) - CSS에서는 별도 처리 불필요 (font metrics로 처리됨)

                            const align = pp['hh:align'];
                            if (align) {
                                // 수평 정렬
                                if (align['@_horizontal']) {
                                    let h = align['@_horizontal'].toLowerCase();
                                    if (h === 'justify') css += `text-align: justify; `;
                                    else if (h === 'center') css += `text-align: center; `;
                                    else if (h === 'right') css += `text-align: right; `;
                                    else if (h === 'distribute' || h === 'distribute_space') css += `text-align: justify; `;
                                    else css += `text-align: left; `;
                                }
                                // 수직 정렬
                                if (align['@_vertical']) {
                                    const v = align['@_vertical'].toLowerCase();
                                    if (v === 'center') css += `vertical-align: middle; `;
                                    else if (v === 'bottom') css += `vertical-align: bottom; `;
                                }
                            }

                            // hp:switch 구조에서 margin/lineSpacing 파싱
                            // hp:case 값 + 표준 HWPUNIT 변환(/100)이 공식 뷰어 PDF와 일치
                            // hp:default 값은 2배 스케일 → CHAR 유닛 fallback 시 /2 보정 필요
                            let marginNode = pp['hh:margin'];
                            let lineSpacing = pp['hh:lineSpacing'];
                            let marginScale = 100; // HWPUNIT → pt 변환 분모
                            const switchNode = pp['hp:switch'];
                            if (switchNode) {
                                const caseNode = switchNode['hp:case'];
                                const defaultNode = switchNode['hp:default'];
                                if (caseNode) {
                                    const caseMargin = caseNode['hh:margin'];
                                    if (caseMargin) {
                                        // CHAR 유닛 사용 여부 확인 — CHAR이면 default로 fallback (2배 보정)
                                        const hasCharUnit = ['hc:left','hc:right','hc:prev','hc:next','hc:intent'].some(
                                            k => caseMargin[k]?.['@_unit'] === 'CHAR'
                                        );
                                        if (hasCharUnit && defaultNode?.['hh:margin']) {
                                            marginNode = defaultNode['hh:margin'];
                                            marginScale = 200; // hp:default 값은 2배 → /200으로 보정
                                        } else {
                                            marginNode = caseMargin;
                                        }
                                    }
                                    if (caseNode['hh:lineSpacing']) lineSpacing = caseNode['hh:lineSpacing'];
                                } else if (defaultNode) {
                                    if (defaultNode['hh:margin']) marginNode = defaultNode['hh:margin'];
                                    if (defaultNode['hh:lineSpacing']) lineSpacing = defaultNode['hh:lineSpacing'];
                                    marginScale = 200; // hp:default만 있으면 2배 보정
                                }
                            }

                            if (lineSpacing && lineSpacing['@_type'] === 'PERCENT') {
                                css += `line-height: ${parseInt(lineSpacing['@_value']) / 100}; `;
                            }

                            // 문단 여백
                            if (marginNode) {
                                const leftNode = marginNode['hc:left'];
                                const rightNode = marginNode['hc:right'];
                                const prevNode = marginNode['hc:prev'];
                                const nextNode = marginNode['hc:next'];
                                const intentNode = marginNode['hc:intent'];

                                if (leftNode) {
                                    const v = parseInt(leftNode['@_value'] || '0');
                                    if (v !== 0) css += `margin-left: ${(v / marginScale).toFixed(1)}pt; `;
                                }
                                if (rightNode) {
                                    const v = parseInt(rightNode['@_value'] || '0');
                                    if (v !== 0) css += `margin-right: ${(v / marginScale).toFixed(1)}pt; `;
                                }
                                if (prevNode) {
                                    const v = parseInt(prevNode['@_value'] || '0');
                                    css += `margin-top: ${(v / marginScale).toFixed(1)}pt; `;
                                } else {
                                    css += `margin-top: 0; `;
                                }
                                if (nextNode) {
                                    const v = parseInt(nextNode['@_value'] || '0');
                                    css += `margin-bottom: ${(v / marginScale).toFixed(1)}pt; `;
                                } else {
                                    css += `margin-bottom: 0; `;
                                }
                                if (intentNode) {
                                    const v = parseInt(intentNode['@_value'] || '0');
                                    if (v !== 0) css += `text-indent: ${(v / marginScale).toFixed(1)}pt; `;
                                }
                            } else {
                                css += `margin-top: 0; margin-bottom: 0; `;
                            }

                            // 문단 테두리 (border)
                            const paraBorder = pp['hh:border'];
                            if (paraBorder) {
                                const bfRef = paraBorder['@_borderFillIDRef'];
                                if (bfRef && this.borderFillMap[bfRef]) {
                                    const bf = this.borderFillMap[bfRef];
                                    if (bf.topBorder.type !== 'NONE') css += `border-top: ${this.borderInfoToCss(bf.topBorder)}; `;
                                    if (bf.bottomBorder.type !== 'NONE') css += `border-bottom: ${this.borderInfoToCss(bf.bottomBorder)}; `;
                                    if (bf.leftBorder.type !== 'NONE') css += `border-left: ${this.borderInfoToCss(bf.leftBorder)}; `;
                                    if (bf.rightBorder.type !== 'NONE') css += `border-right: ${this.borderInfoToCss(bf.rightBorder)}; `;
                                    const oL = parseInt(paraBorder['@_offsetLeft'] || '0') / 283.465;
                                    const oR = parseInt(paraBorder['@_offsetRight'] || '0') / 283.465;
                                    const oT = parseInt(paraBorder['@_offsetTop'] || '0') / 283.465;
                                    const oB = parseInt(paraBorder['@_offsetBottom'] || '0') / 283.465;
                                    if (oL || oR || oT || oB) {
                                        css += `padding: ${oT}mm ${oR}mm ${oB}mm ${oL}mm; `;
                                    }
                                }
                            }

                            // 글머리 기호 (heading type=BULLET) — paraPr margin이 들여쓰기 처리
                            const heading = pp['hh:heading'];
                            if (heading && heading['@_type'] === 'BULLET') {
                                const idRef = heading['@_idRef'] || '1';
                                const bulletInfo = bulletMap[idRef];
                                const bulletChar = bulletInfo?.char || '•';
                                const autoIndent = bulletInfo?.autoIndent ?? false;
                                // autoIndent: 글머리 기호 뒤 텍스트 위치에 후속 줄 정렬 (hanging indent)
                                if (autoIndent) {
                                    css += `padding-left: 1.2em; text-indent: -1.2em; `;
                                }
                                result.css += `.para-${id} { ${css} }\n`;
                                result.css += `.para-${id}::before { content: '${bulletChar === '-' ? '- ' : bulletChar + '\\00a0'}'; }\n`;
                            }
                            // 번호 매기기 (heading type=NUMBER) — paraPr margin이 들여쓰기 처리
                            else if (heading && heading['@_type'] === 'NUMBER') {
                                const idRef = heading['@_idRef'] || '1';
                                const level = parseInt(heading['@_level'] || '0') + 1;
                                const counterName = `hwpx-num-${idRef}-${level}`;
                                const levelInfo = this.numberingMap[idRef]?.[String(level)];
                                let counterStyle = 'decimal';
                                if (levelInfo) {
                                    switch (levelInfo.numFormat) {
                                        case 'HANGUL_SYLLABLE': counterStyle = 'korean-hangul-formal'; break;
                                        case 'LATIN_UPPER': counterStyle = 'upper-latin'; break;
                                        case 'LATIN_LOWER': counterStyle = 'lower-latin'; break;
                                        case 'ROMAN_UPPER': counterStyle = 'upper-roman'; break;
                                        case 'ROMAN_LOWER': counterStyle = 'lower-roman'; break;
                                        default: counterStyle = 'decimal';
                                    }
                                    // autoIndent: 번호 뒤 텍스트 위치에 후속 줄 정렬
                                    if (levelInfo.autoIndent) {
                                        css += `padding-left: 1.6em; text-indent: -1.6em; `;
                                    }
                                }
                                result.css += `.para-${id} { ${css} counter-increment: ${counterName}; }\n`;
                                result.css += `.para-${id}::before { content: counter(${counterName}, ${counterStyle}) '. '; }\n`;
                            }
                            else {
                                result.css += `.para-${id} { ${css} }\n`;
                            }
                        }
                    }
                    // ===== 스타일 정의 (hh:styles) =====
                    const styles = head['hh:styles'];
                    if (styles) {
                        const styleList = Array.isArray(styles['hh:style']) ? styles['hh:style'] : (styles['hh:style'] ? [styles['hh:style']] : []);
                        for (const st of styleList) {
                            const sid = String(st['@_id'] || '');
                            this.styleMap[sid] = {
                                paraPrIDRef: String(st['@_paraPrIDRef'] || ''),
                                charPrIDRef: String(st['@_charPrIDRef'] || ''),
                                type: st['@_type'] || 'PARA',
                            };
                        }
                    }
                }
            }
        } catch (e) {
            console.error('Error parsing header.xml', e);
        }

        // 3. 본문(sectionN.xml) 파싱
        let sectionIndex = 0;
        let htmlParts: string[] = [];

        while (true) {
            const sectionPath = `Contents/section${sectionIndex}.xml`;
            const sectionXml = await zip.file(sectionPath)?.async('string');
            if (!sectionXml) break;

            try {
                const processedSectionXml = sectionXml
                    .replace(/<hp:lineBreak\s*\/>|<hp:lineBreak\s*><\/hp:lineBreak>/g, '\n')
                    .replace(/<hp:tab[^\/]*\/>|<hp:tab[^>]*><\/hp:tab>/g, '\t')
                    .replace(/<hp:nbSpace\s*\/>|<hp:nbSpace\s*><\/hp:nbSpace>/g, '\u00A0')
                    .replace(/<hp:fwSpace\s*\/>|<hp:fwSpace\s*><\/hp:fwSpace>/g, '\u3000')
                    .replace(/<hp:hyphen\s*\/>|<hp:hyphen\s*><\/hp:hyphen>/g, '\u00AD')
                    .replace(/<hp:bookmark\b[^>]*\/>|<hp:bookmark\b[^>]*><\/hp:bookmark>/g, '')
                    .replace(/<hp:titleMark\b[^>]*\/>|<hp:titleMark\b[^>]*><\/hp:titleMark>/g, '')
                    .replace(/<hp:markpenBegin\b[^>]*?color="([^"]*)"[^>]*?\/?>/g, '\uE000$1\uE001')
                    .replace(/<hp:markpenEnd\s*\/?>/g, '\uE002')
                    .replace(/<hp:pageNum\b[^>]*\/>/g, '\uE003')
                    .replace(/<hp:insertBegin\b[^>]*\/>/g, '')
                    .replace(/<hp:insertEnd\b[^>]*\/>/g, '')
                    .replace(/<hp:deleteBegin\b[^>]*\/>/g, '')
                    .replace(/<hp:deleteEnd\b[^>]*\/>/g, '');
                const sectionObj = parser.parse(processedSectionXml);

                let pageCss = defaultPageCss;
                let pageDataAttrs = '';

                let headerHtml = '';
                let footerHtml = '';
                let headerMm = 0;
                let footerMm = 0;
                let hideFirstHeader = false;
                let hideFirstFooter = false;

                const secPrNode = this.findDeepNode(sectionObj, 'hp:secPr');
                if (secPrNode && secPrNode['hp:pagePr']) {
                    const pagePr = secPrNode['hp:pagePr'];
                    let widthMm = parseInt(pagePr['@_width'] || '59528') / 283.465;
                    let heightMm = parseInt(pagePr['@_height'] || '84188') / 283.465;

                    const landscapeVal = String(pagePr['@_landscape'] || '').trim().toLowerCase();
                    // landscape 속성이 있으면 (NARROWLY/WIDELY 등) 가로 방향 — 실제 레이아웃 크기 우선
                    const isLandscape = landscapeVal !== '' && landscapeVal !== '0' && landscapeVal !== 'false' && landscapeVal !== 'undefined';
                    if (isLandscape && widthMm < heightMm) {
                        const temp = widthMm;
                        widthMm = heightMm;
                        heightMm = temp;
                    }

                    let paddingCss = '';
                    const margin = pagePr['hp:margin'] || pagePr['hp:pageMar'];
                    if (margin) {
                        headerMm = parseInt(margin['@_header'] || '0') / 283.465;
                        footerMm = parseInt(margin['@_footer'] || '0') / 283.465;
                        // top+header, bottom+footer = 본문 영역까지의 전체 여백
                        const topMm = (parseInt(margin['@_top'] || '0') + parseInt(margin['@_header'] || '0')) / 283.465;
                        const bottomMm = (parseInt(margin['@_bottom'] || '0') + parseInt(margin['@_footer'] || '0')) / 283.465;
                        const leftMm = parseInt(margin['@_left'] || '0') / 283.465;
                        const rightMm = parseInt(margin['@_right'] || '0') / 283.465;
                        paddingCss = `padding: ${topMm}mm ${rightMm}mm ${bottomMm}mm ${leftMm}mm; `;
                    } else {
                        paddingCss = `padding: 20mm; `;
                    }
                    pageCss = `width: ${widthMm}mm; min-height: ${heightMm}mm; ${paddingCss} box-sizing: border-box; position: relative; overflow: hidden;`;
                    pageDataAttrs = `data-width="${widthMm}" data-height="${heightMm}"`;

                    // 페이지 테두리/채움 (pageBorderFill)
                    const pageBorderFills = secPrNode['hp:pageBorderFill'];
                    if (pageBorderFills) {
                        const pbfList = Array.isArray(pageBorderFills) ? pageBorderFills : [pageBorderFills];
                        const pbf = pbfList.find((p: any) => p['@_type'] === 'BOTH') || pbfList[0];
                        if (pbf) {
                            const bfRef = pbf['@_borderFillIDRef'];
                            if (bfRef && this.borderFillMap[bfRef]) {
                                const bf = this.borderFillMap[bfRef];
                                let pageBorderCss = '';
                                if (bf.topBorder.type !== 'NONE') pageBorderCss += `border-top: ${this.borderInfoToCss(bf.topBorder)}; `;
                                if (bf.bottomBorder.type !== 'NONE') pageBorderCss += `border-bottom: ${this.borderInfoToCss(bf.bottomBorder)}; `;
                                if (bf.leftBorder.type !== 'NONE') pageBorderCss += `border-left: ${this.borderInfoToCss(bf.leftBorder)}; `;
                                if (bf.rightBorder.type !== 'NONE') pageBorderCss += `border-right: ${this.borderInfoToCss(bf.rightBorder)}; `;
                                if (bf.fillCss) {
                                    pageBorderCss += bf.fillCss + ' ';
                                } else if (bf.faceColor && bf.faceColor !== 'none') {
                                    pageBorderCss += `background-color: ${bf.faceColor}; `;
                                }
                                if (pageBorderCss) pageCss += ` ${pageBorderCss}`;
                            }
                        }
                    }

                    // 페이지 표시 설정 (visibility)
                    const visibility = secPrNode['hp:visibility'];
                    if (visibility) {
                        hideFirstHeader = visibility['@_hideFirstHeader'] === '1' || visibility['@_hideFirstHeader'] === 'true';
                        hideFirstFooter = visibility['@_hideFirstFooter'] === '1' || visibility['@_hideFirstFooter'] === 'true';
                    }
                }

                // 머리글 (header) 파싱
                if (secPrNode) {
                    const headers = secPrNode['hp:header'];
                    if (headers) {
                        const headerList = Array.isArray(headers) ? headers : [headers];
                        const headerNode = headerList.find((h: any) => h['@_applyPageType'] === 'BOTH') || headerList[0];
                        if (headerNode) {
                            const subList = headerNode['hp:subList'];
                            if (subList) {
                                headerHtml = this.extractHtml(subList, imageMap);
                            }
                        }
                    }
                    // 바닥글 (footer) 파싱
                    const footers = secPrNode['hp:footer'];
                    if (footers) {
                        const footerList = Array.isArray(footers) ? footers : [footers];
                        const footerNode = footerList.find((f: any) => f['@_applyPageType'] === 'BOTH') || footerList[0];
                        if (footerNode) {
                            const subList = footerNode['hp:subList'];
                            if (subList) {
                                footerHtml = this.extractHtml(subList, imageMap);
                            }
                        }
                    }
                }

                // 실제 XML 루트 태그를 경로에 반영 (hs:sec, hp:sec 등)
                const rootTag = Object.keys(sectionObj).find(k => !k.startsWith('@_') && !k.startsWith('#') && !k.startsWith('?')) || 'hs:sec';
                const sectionXmlPath = `/${rootTag}`;
                const rootObj = sectionObj[rootTag] || sectionObj;
                const sectionHtml = this.extractHtml(rootObj, imageMap, sectionXmlPath);
                const headerDataAttr = headerHtml ? ` data-header="${encodeURIComponent(headerHtml)}"` : '';
                const footerDataAttr = footerHtml ? ` data-footer="${encodeURIComponent(footerHtml)}"` : '';
                const headerPosAttr = headerMm ? ` data-header-top="${headerMm.toFixed(1)}"` : '';
                const footerPosAttr = footerMm ? ` data-footer-bottom="${footerMm.toFixed(1)}"` : '';
                const hideFirstAttr = (hideFirstHeader ? ' data-hide-first-header="1"' : '') + (hideFirstFooter ? ' data-hide-first-footer="1"' : '');
                htmlParts.push(`
                    <div class="hwpx-section-container" data-hwpx-file="${sectionPath}" style="display: none;" ${pageDataAttrs}${headerDataAttr}${footerDataAttr}${headerPosAttr}${footerPosAttr}${hideFirstAttr}>
                        <div class="hwpx-page-template" style="${pageCss}">${sectionHtml}</div>
                    </div>
                `);
            } catch (e) {
                console.error(`Error parsing ${sectionPath}`, e);
            }

            sectionIndex++;
        }

        result.html = htmlParts.join('\n');

        // 브라우저 렌더링 시점에 각 요소의 위치를 측정하여 페이지 분할하는 JS (표는 overflow clip 방식으로 행 중간에서도 분할)
        result.html += `
        <div id="hwpx-render-root"></div>
        <script>
            (function renderPagination() {
                const root = document.getElementById('hwpx-render-root');
                const sections = document.querySelectorAll('.hwpx-section-container');
                const mmToPx = (mm) => mm * 3.779527;

                // 모든 이미지가 로드된 후 페이지네이션 실행
                var allImgs = document.querySelectorAll('.hwpx-section-container img');
                var imgPromises = [];
                for (var ii = 0; ii < allImgs.length; ii++) {
                    (function(img) {
                        if (!img.complete) {
                            imgPromises.push(new Promise(function(resolve) {
                                img.onload = resolve;
                                img.onerror = resolve;
                            }));
                        }
                    })(allImgs[ii]);
                }
                Promise.all(imgPromises).then(function() { doPagination(); });

                function doPagination() {

                sections.forEach(function(section) {
                    var template = section.querySelector('.hwpx-page-template');

                    // 머리글/바닥글 데이터
                    var rawHeader = section.getAttribute('data-header');
                    var rawFooter = section.getAttribute('data-footer');
                    var headerContent = rawHeader ? decodeURIComponent(rawHeader) : '';
                    var footerContent = rawFooter ? decodeURIComponent(rawFooter) : '';
                    var headerTopMm = parseFloat(section.getAttribute('data-header-top') || '0');
                    var footerBottomMm = parseFloat(section.getAttribute('data-footer-bottom') || '0');
                    var hideFirstHeader = section.getAttribute('data-hide-first-header') === '1';
                    var hideFirstFooter = section.getAttribute('data-hide-first-footer') === '1';

                    template.style.display = 'block';
                    template.style.height = 'auto';
                    template.style.minHeight = 'auto';
                    template.style.overflow = 'visible';
                    template.style.visibility = 'hidden';
                    template.style.position = 'absolute';
                    template.style.top = '0';
                    template.style.left = '-99999px';
                    document.body.appendChild(template);

                    const widthMm = parseFloat(section.getAttribute('data-width') || 210);
                    const heightMm = parseFloat(section.getAttribute('data-height') || 297);

                    const cs = window.getComputedStyle(template);
                    const padTop = parseFloat(cs.paddingTop);
                    const padBottom = parseFloat(cs.paddingBottom);
                    const padLeft = parseFloat(cs.paddingLeft);
                    const padRight = parseFloat(cs.paddingRight);
                    // 브라우저의 실제 mm→px 비율로 정확한 페이지 높이 계산
                    const ruler = document.createElement('div');
                    ruler.style.cssText = 'position:absolute;visibility:hidden;width:0;height:' + heightMm + 'mm;';
                    document.body.appendChild(ruler);
                    const pageHeightPx = ruler.getBoundingClientRect().height;
                    document.body.removeChild(ruler);
                    const contentMaxH = pageHeightPx - padTop - padBottom;

                    document.body.removeChild(template);

                    // 페이지 높이를 px로 통일 (mm-to-px 변환 차이로 인한 클리핑 방지)
                    const pageHeightForLayout = padTop + contentMaxH + padBottom;
                    function newPage() {
                        const p = document.createElement('div');
                        p.className = 'hwpx-page';
                        p.style.width = widthMm + 'mm';
                        p.style.height = pageHeightForLayout + 'px';
                        p.style.minHeight = pageHeightForLayout + 'px';
                        p.style.boxSizing = 'border-box';
                        p.style.position = 'relative';
                        p.style.overflow = 'hidden';
                        p.style.padding = padTop + 'px ' + padRight + 'px ' + padBottom + 'px ' + padLeft + 'px';
                        return p;
                    }

                    const measure = document.createElement('div');
                    measure.style.cssText = template.style.cssText;
                    measure.style.position = 'absolute';
                    measure.style.top = '-99999px';
                    measure.style.left = '-99999px';
                    measure.style.height = 'auto';
                    measure.style.minHeight = 'auto';
                    measure.style.overflow = 'visible';
                    measure.style.visibility = 'hidden';
                    measure.style.padding = padTop + 'px ' + padRight + 'px ' + padBottom + 'px ' + padLeft + 'px';
                    document.body.appendChild(measure);

                    function getHeight() { return measure.scrollHeight - padTop - padBottom; }

                    let pages = [];
                    let curPage = newPage();
                    let curMaxH = contentMaxH;
                    measure.innerHTML = '';

                    function finalizePage() {
                        if (curPage.children.length > 0) pages.push(curPage);
                        curPage = newPage();
                        measure.innerHTML = '';
                        curMaxH = contentMaxH;
                    }

                    const children = Array.from(template.children);

                    for (let ci = 0; ci < children.length; ci++) {
                        const child = children[ci];

                        if (child.tagName === 'TABLE') {
                            // 테이블 분할: 실시간 측정 + 행/셀 내용 DOM 분리
                            var hasRepeatHeader = child.getAttribute('data-repeat-header') === '1';
                            var origThead = hasRepeatHeader ? child.querySelector('thead') : null;
                            var origTbody = child.querySelector('tbody');
                            var bodyRows = origTbody
                                ? Array.from(origTbody.querySelectorAll(':scope > tr'))
                                : Array.from(child.querySelectorAll(':scope > tr')).filter(function(tr) {
                                    return !origThead || !origThead.contains(tr);
                                });

                            // 전체 테이블 높이를 실시간 측정
                            var fullClone = child.cloneNode(true);
                            fullClone.style.margin = '0';
                            measure.appendChild(fullClone);
                            var totalTableH = fullClone.getBoundingClientRect().height;
                            measure.removeChild(fullClone);

                            var usedH = getHeight();
                            var availH = curMaxH - usedH;

                            if (totalTableH <= availH) {
                                curPage.appendChild(child.cloneNode(true));
                                measure.appendChild(child.cloneNode(true));
                            } else {
                                if (availH < 30 && curPage.children.length > 0) {
                                    pages.push(curPage);
                                    curPage = newPage();
                                    measure.innerHTML = '';
                                    curMaxH = contentMaxH;
                                    usedH = 0;
                                    availH = contentMaxH;
                                }

                                // 행의 셀 내용을 높이 기준으로 픽셀 단위 overflow clip 분리
                                function splitRowAtHeight(tblRef, thRef, rowNode, maxH) {
                                    var cells = Array.from(rowNode.querySelectorAll(':scope > td, :scope > th'));

                                    // 셀 패딩 측정 (정확한 clip 높이 계산용)
                                    var tmpTbl = tblRef.cloneNode(false);
                                    tmpTbl.style.margin = '0';
                                    if (thRef) tmpTbl.appendChild(thRef.cloneNode(true));
                                    var tmpTb = document.createElement('tbody');
                                    var tmpTr = rowNode.cloneNode(true);
                                    tmpTb.appendChild(tmpTr);
                                    tmpTbl.appendChild(tmpTb);
                                    measure.appendChild(tmpTbl);

                                    var tmpCells = Array.from(tmpTr.querySelectorAll(':scope > td, :scope > th'));
                                    var cellPads = [];
                                    for (var ci2 = 0; ci2 < tmpCells.length; ci2++) {
                                        var cs = getComputedStyle(tmpCells[ci2]);
                                        cellPads.push({
                                            pt: parseFloat(cs.paddingTop) || 0,
                                            pb: parseFloat(cs.paddingBottom) || 0
                                        });
                                    }
                                    measure.removeChild(tmpTbl);

                                    // top (클리핑) / bottom (나머지) 행 생성
                                    var topTr = rowNode.cloneNode(false);
                                    var bottomTr = rowNode.cloneNode(false);

                                    for (var ci3 = 0; ci3 < cells.length; ci3++) {
                                        var topCell = cells[ci3].cloneNode(false);
                                        var bottomCell = cells[ci3].cloneNode(false);
                                        topCell.style.height = 'auto';
                                        bottomCell.style.height = 'auto';

                                        var contentClipH = maxH - cellPads[ci3].pt - cellPads[ci3].pb;
                                        if (contentClipH < 1) contentClipH = 1;

                                        var kids = Array.from(cells[ci3].childNodes);
                                        if (kids.length === 0) {
                                            topTr.appendChild(topCell);
                                            bottomTr.appendChild(bottomCell);
                                            continue;
                                        }

                                        // Top: overflow hidden으로 contentClipH만큼만 표시
                                        var topWrap = document.createElement('div');
                                        topWrap.style.overflow = 'hidden';
                                        topWrap.style.maxHeight = contentClipH + 'px';
                                        for (var k = 0; k < kids.length; k++) {
                                            topWrap.appendChild(kids[k].cloneNode(true));
                                        }
                                        topCell.appendChild(topWrap);

                                        // Bottom: 이미 표시된 부분을 negative margin으로 건너뛰기
                                        var botOuter = document.createElement('div');
                                        botOuter.style.overflow = 'hidden';
                                        var botInner = document.createElement('div');
                                        botInner.style.marginTop = '-' + contentClipH + 'px';
                                        for (var k2 = 0; k2 < kids.length; k2++) {
                                            botInner.appendChild(kids[k2].cloneNode(true));
                                        }
                                        botOuter.appendChild(botInner);
                                        bottomCell.appendChild(botOuter);

                                        topTr.appendChild(topCell);
                                        bottomTr.appendChild(bottomCell);
                                    }
                                    return { topRow: topTr, bottomRow: bottomTr };
                                }

                                var rowIdx = 0;
                                var isFirstTablePage = true;
                                var pendingBottomRow = null;
                                var safety = 0;

                                while ((rowIdx < bodyRows.length || pendingBottomRow) && safety < 200) {
                                    safety++;
                                    var pageAvail = isFirstTablePage ? availH : contentMaxH;

                                    // 실시간 측정용 테이블 생성 (measure에 추가)
                                    var liveTbl = child.cloneNode(false);
                                    liveTbl.style.margin = '0';
                                    if (origThead) liveTbl.appendChild(origThead.cloneNode(true));
                                    var liveTb = document.createElement('tbody');
                                    liveTbl.appendChild(liveTb);
                                    measure.appendChild(liveTbl);

                                    var baseH = liveTbl.getBoundingClientRect().height; // thead 높이
                                    var rowsAdded = 0;

                                    // pending 행이 있으면 먼저 추가
                                    if (pendingBottomRow) {
                                        liveTb.appendChild(pendingBottomRow.cloneNode(true));
                                        var newH = liveTbl.getBoundingClientRect().height;
                                        if (newH <= pageAvail + 0.5) {
                                            rowsAdded++;
                                            pendingBottomRow = null;
                                            baseH = newH;
                                        } else {
                                            // pending 행도 안 맞음 → 다시 분할
                                            liveTb.removeChild(liveTb.lastChild);
                                            var space2 = pageAvail - baseH;
                                            if (space2 > 10) {
                                                var sp2 = splitRowAtHeight(child, origThead, pendingBottomRow, space2);
                                                liveTb.appendChild(sp2.topRow);
                                                pendingBottomRow = sp2.bottomRow;
                                                rowsAdded++;
                                            }
                                            // 이 페이지는 여기까지
                                            var pageTblClone = liveTbl.cloneNode(true);
                                            measure.removeChild(liveTbl);
                                            curPage.appendChild(pageTblClone);
                                            pages.push(curPage);
                                            curPage = newPage();
                                            measure.innerHTML = '';
                                            curMaxH = contentMaxH;
                                            isFirstTablePage = false;
                                            continue;
                                        }
                                    }

                                    // 행을 하나씩 추가하며 실시간 측정
                                    while (rowIdx < bodyRows.length) {
                                        liveTb.appendChild(bodyRows[rowIdx].cloneNode(true));
                                        var curH = liveTbl.getBoundingClientRect().height;
                                        if (curH > pageAvail + 0.5) {
                                            // 이 행이 overflow → 제거 후 셀 내용 분리
                                            liveTb.removeChild(liveTb.lastChild);
                                            var spaceLeft = pageAvail - baseH;
                                            if (spaceLeft > 10) {
                                                var sp = splitRowAtHeight(child, origThead, bodyRows[rowIdx], spaceLeft);
                                                liveTb.appendChild(sp.topRow);
                                                pendingBottomRow = sp.bottomRow;
                                                rowIdx++;
                                            }
                                            break;
                                        }
                                        baseH = curH;
                                        rowsAdded++;
                                        rowIdx++;
                                    }

                                    // 이 페이지의 테이블을 curPage에 추가
                                    var pageTbl = liveTbl.cloneNode(true);
                                    measure.removeChild(liveTbl);

                                    if (rowsAdded > 0 || liveTb.childNodes.length > 0) {
                                        curPage.appendChild(pageTbl);
                                    }

                                    if (rowIdx < bodyRows.length || pendingBottomRow) {
                                        pages.push(curPage);
                                        curPage = newPage();
                                        measure.innerHTML = '';
                                        curMaxH = contentMaxH;
                                    }

                                    isFirstTablePage = false;
                                }

                                // measure placeholder
                                var lastTbl = curPage.querySelector('table');
                                if (lastTbl) measure.appendChild(lastTbl.cloneNode(true));
                            }

                        } else {
                            const testDiv = measure.cloneNode(true);
                            testDiv.style.cssText = measure.style.cssText;
                            document.body.appendChild(testDiv);
                            testDiv.appendChild(child.cloneNode(true));
                            const h = testDiv.scrollHeight - padTop - padBottom;
                            document.body.removeChild(testDiv);

                            if (h > curMaxH && curPage.children.length > 0) {
                                finalizePage();
                            }
                            curPage.appendChild(child.cloneNode(true));
                            measure.appendChild(child.cloneNode(true));
                        }
                    }

                    if (curPage.children.length > 0) pages.push(curPage);

                    pages.forEach(function(p, pi) {
                        // 머리글 삽입
                        if (headerContent && !(hideFirstHeader && pi === 0)) {
                            var hdr = document.createElement('div');
                            hdr.className = 'hwpx-header';
                            hdr.style.cssText = 'position: absolute; top: ' + mmToPx(headerTopMm) + 'px; left: ' + padLeft + 'px; right: ' + padRight + 'px; font-size: 10pt;';
                            hdr.innerHTML = headerContent;
                            p.appendChild(hdr);
                        }
                        // 바닥글 삽입
                        if (footerContent && !(hideFirstFooter && pi === 0)) {
                            var ftr = document.createElement('div');
                            ftr.className = 'hwpx-footer';
                            ftr.style.cssText = 'position: absolute; bottom: ' + mmToPx(footerBottomMm) + 'px; left: ' + padLeft + 'px; right: ' + padRight + 'px; font-size: 10pt;';
                            ftr.innerHTML = footerContent;
                            p.appendChild(ftr);
                        }
                        root.appendChild(p);
                        // 페이지 번호 삽입
                        p.querySelectorAll('.hwpx-pagenum').forEach(function(el) {
                            el.textContent = String(pi + 1);
                        });
                        // 각주 수집 및 페이지 하단에 배치
                        var fnContents = p.querySelectorAll('.hwpx-footnote-content');
                        if (fnContents.length > 0) {
                            var fnArea = document.createElement('div');
                            fnArea.className = 'hwpx-footnote-area';
                            fnArea.style.cssText = 'position: absolute; bottom: ' + padBottom + 'px; left: ' + padLeft + 'px; right: ' + padRight + 'px; border-top: 1px solid #000; padding-top: 4px; font-size: 9pt;';
                            fnContents.forEach(function(fc) {
                                var fnNum = fc.getAttribute('data-fn-num') || '';
                                var fnDiv = document.createElement('div');
                                fnDiv.innerHTML = '<sup>' + fnNum + ')</sup> ' + fc.innerHTML;
                                fnArea.appendChild(fnDiv);
                                fc.remove();
                            });
                            p.appendChild(fnArea);
                        }
                    });
                    section.remove();
                    document.body.removeChild(measure);
                });
                // 미주 수집 — 문서 끝에 모아서 표시
                var enContents = root.querySelectorAll('.hwpx-endnote-content');
                if (enContents.length > 0) {
                    var enArea = document.createElement('div');
                    enArea.className = 'hwpx-endnote-area';
                    enArea.style.cssText = 'border-top: 2px solid #000; padding-top: 8px; margin-top: 16px; font-size: 9pt;';
                    enContents.forEach(function(ec) {
                        var enNum = ec.getAttribute('data-en-num') || '';
                        var enDiv = document.createElement('div');
                        enDiv.innerHTML = '<sup>' + enNum + ')</sup> ' + ec.innerHTML;
                        enArea.appendChild(enDiv);
                        ec.remove();
                    });
                    root.appendChild(enArea);
                }
                } // end doPagination
            })();
        </script>
        `;
        return result;
    }

    /**
     * JSON으로 변환된 XML 트리에서 재귀적으로 요소(단락, 표, 이미지 등)를 탐색하여 HTML로 변환합니다.
     */
    private static extractHtml(obj: any, imageMap: Record<string, string>, xmlPath: string = ''): string {
        let html = '';

        if (Array.isArray(obj)) {
            for (let ai = 0; ai < obj.length; ai++) {
                html += this.extractHtml(obj[ai], imageMap, xmlPath);
            }
        } else if (typeof obj === 'object' && obj !== null) {

            // 표 파싱 (hp:tbl)
            if (obj['hp:tbl']) {
                const tables = Array.isArray(obj['hp:tbl']) ? obj['hp:tbl'] : [obj['hp:tbl']];
                for (let ti = 0; ti < tables.length; ti++) {
                    const t = tables[ti];
                    const tblPath = xmlPath + (tables.length > 1 ? `/hp:tbl[${ti}]` : '/hp:tbl');
                    // 테이블 너비: 항상 콘텐츠 영역에 맞춤 (table-layout: fixed + 셀 비율)
                    const tableWidthHwp = (t['hp:sz'] && t['hp:sz']['@_width']) ? parseInt(t['hp:sz']['@_width']) : 0;
                    let tableStyle = 'table-layout: fixed;';

                    const repeatHeader = t['@_repeatHeader'] === '1' || t['@_repeatHeader'] === 'true';
                    const dataRepeat = repeatHeader ? ' data-repeat-header="1"' : '';

                    // 표 외부 여백 (hp:outMargin) — margin + width를 함께 조정하여 overflow 방지
                    const outMargin = t['hp:outMargin'];
                    let tableMargin = 'margin: 0;';
                    let tableWidth = 'width: 100%;';
                    if (outMargin) {
                        const omt = (parseInt(outMargin['@_top'] || '0') / 283.465).toFixed(1);
                        const omb = (parseInt(outMargin['@_bottom'] || '0') / 283.465).toFixed(1);
                        const oml = (parseInt(outMargin['@_left'] || '0') / 283.465).toFixed(1);
                        const omr = (parseInt(outMargin['@_right'] || '0') / 283.465).toFixed(1);
                        tableMargin = `margin: ${omt}mm ${omr}mm ${omb}mm ${oml}mm;`;
                        const totalHoriz = parseFloat(oml) + parseFloat(omr);
                        if (totalHoriz > 0) {
                            tableWidth = `width: calc(100% - ${totalHoriz.toFixed(1)}mm);`;
                        }
                    }

                    // 셀 간격 (cellSpacing)
                    const cellSpacing = parseInt(t['@_cellSpacing'] || '0');
                    if (cellSpacing > 0) {
                        const spaceMm = (cellSpacing / 283.465).toFixed(1);
                        tableStyle += ` border-spacing: ${spaceMm}mm;`;
                        tableStyle = tableStyle.replace('border-collapse: collapse;', 'border-collapse: separate;');
                    }

                    // 표 기본 셀 여백 (hp:inMargin → 개별 셀에 cellMargin이 없을 때 사용)
                    let defaultCellPadding = 'padding: 0;';
                    const inMargin = t['hp:inMargin'];
                    if (inMargin) {
                        const imt = (parseInt(inMargin['@_top'] || '0') / 283.465).toFixed(1);
                        const imb = (parseInt(inMargin['@_bottom'] || '0') / 283.465).toFixed(1);
                        const iml = (parseInt(inMargin['@_left'] || '0') / 283.465).toFixed(1);
                        const imr = (parseInt(inMargin['@_right'] || '0') / 283.465).toFixed(1);
                        defaultCellPadding = `padding: ${imt}mm ${imr}mm ${imb}mm ${iml}mm;`;
                    }

                    html += `<table border="0"${dataRepeat} data-hwpx="${tblPath}" style="border-collapse: collapse; ${tableMargin} ${tableWidth} ${tableStyle}">`;
                    const trs = Array.isArray(t['hp:tr']) ? t['hp:tr'] : (t['hp:tr'] ? [t['hp:tr']] : []);
                    for (let ri = 0; ri < trs.length; ri++) {
                        const tr = trs[ri];
                        const trPath = `${tblPath}/hp:tr[${ri}]`;
                        if (ri === 0 && repeatHeader) html += `<thead>`;
                        html += `<tr data-hwpx="${trPath}">`;
                        const tcs = Array.isArray(tr['hp:tc']) ? tr['hp:tc'] : (tr['hp:tc'] ? [tr['hp:tc']] : []);
                        for (let ci = 0; ci < tcs.length; ci++) {
                            const tc = tcs[ci];
                            const tcPath = `${trPath}/hp:tc[${ci}]`;
                            const colSpan = tc['@_colSpan'] || tc['hp:cellSpan']?.['@_colSpan'] || 1;
                            const rowSpan = tc['@_rowSpan'] || tc['hp:cellSpan']?.['@_rowSpan'] || 1;

                            let cellStyle = '';
                            if (tc['hp:cellSz'] && tc['hp:cellSz']['@_width'] && tableWidthHwp > 0) {
                                const cellWidthHwp = parseInt(tc['hp:cellSz']['@_width']);
                                const pct = (cellWidthHwp / tableWidthHwp * 100).toFixed(2);
                                cellStyle = `width: ${pct}%;`;
                            }

                            let cellPadding = defaultCellPadding;
                            const hasCustomMargin = tc['@_hasMargin'] === '1' || tc['@_hasMargin'] === 'true';
                            const cm = hasCustomMargin ? tc['hp:cellMargin'] : null;
                            if (cm) {
                                const cmt = (parseInt(cm['@_top'] || '0') / 283.465).toFixed(1);
                                const cmb = (parseInt(cm['@_bottom'] || '0') / 283.465).toFixed(1);
                                const cml = (parseInt(cm['@_left'] || '0') / 283.465).toFixed(1);
                                const cmr = (parseInt(cm['@_right'] || '0') / 283.465).toFixed(1);
                                cellPadding = `padding: ${cmt}mm ${cmr}mm ${cmb}mm ${cml}mm;`;
                            }

                            let cellHeightStyle = '';
                            if (tc['hp:cellSz'] && tc['hp:cellSz']['@_height']) {
                                const h = parseInt(tc['hp:cellSz']['@_height']) / 283.465;
                                cellHeightStyle = `height: ${h}mm;`;
                            }

                            // 셀 테두리 + 배경색 (borderFillMap에서 개별 스타일 적용)
                            let borderCss = 'border: 0.06em solid #000;';
                            let bgColor = '';
                            const bfRef = tc['@_borderFillIDRef'];
                            if (bfRef && this.borderFillMap[bfRef]) {
                                const bf = this.borderFillMap[bfRef];
                                borderCss = `border-top: ${this.borderInfoToCss(bf.topBorder)}; border-bottom: ${this.borderInfoToCss(bf.bottomBorder)}; border-left: ${this.borderInfoToCss(bf.leftBorder)}; border-right: ${this.borderInfoToCss(bf.rightBorder)};`;
                                if (bf.fillCss) {
                                    bgColor = bf.fillCss;
                                } else if (bf.faceColor && bf.faceColor !== 'none') {
                                    bgColor = `background-color: ${bf.faceColor};`;
                                }
                            }

                            // 셀 수직 정렬 (hp:subList vertAlign)
                            const subList = tc['hp:subList'];
                            let vAlign = 'top';
                            if (subList && subList['@_vertAlign']) {
                                const va = subList['@_vertAlign'].toUpperCase();
                                if (va === 'CENTER') vAlign = 'middle';
                                else if (va === 'BOTTOM') vAlign = 'bottom';
                                else vAlign = 'top';
                            }

                            html += `<td colspan="${colSpan}" rowspan="${rowSpan}" data-hwpx="${tcPath}" style="${cellPadding} ${borderCss} word-break: break-word; vertical-align: ${vAlign}; ${cellStyle} ${cellHeightStyle} ${bgColor}">`;
                            if (subList) {
                                html += this.extractHtml(subList, imageMap, `${tcPath}/hp:subList`);
                            } else {
                                html += this.extractHtml(tc, imageMap, tcPath);
                            }
                            html += `</td>`;
                        }
                        html += `</tr>`;
                        if (ri === 0 && repeatHeader) html += `</thead><tbody>`;
                    }
                    if (repeatHeader && trs.length > 0) html += `</tbody>`;
                    html += `</table>`;
                }
            }
            // 문단 파싱 (hp:p)
            else if (obj['hp:p']) {
                const paragraphs = Array.isArray(obj['hp:p']) ? obj['hp:p'] : [obj['hp:p']];
                for (let pi = 0; pi < paragraphs.length; pi++) {
                    const p = paragraphs[pi];
                    const pPath = `${xmlPath}/hp:p[${pi}]`;
                    const paraId = p['@_paraPrIDRef'] !== undefined ? p['@_paraPrIDRef'] : '';
                    let paraClass = paraId !== '' ? `para-${paraId}` : '';

                    // 스타일 정의에서 기본 charPr 클래스 가져오기 (run 레벨에서 charPrIDRef가 없을 때 폴백)
                    const styleId = p['@_styleIDRef'];
                    let styleCharClass = '';
                    if (styleId !== undefined && this.styleMap[String(styleId)]) {
                        const style = this.styleMap[String(styleId)];
                        if (style.charPrIDRef) styleCharClass = `char-${style.charPrIDRef}`;
                        // paraPrIDRef가 없으면 스타일의 것 사용
                        if (paraId === '' && style.paraPrIDRef) paraClass = `para-${style.paraPrIDRef}`;
                    }

                    const runs = Array.isArray(p['hp:run']) ? p['hp:run'] : (p['hp:run'] ? [p['hp:run']] : []);
                    const hasTbl = runs.some((r: any) => !!r['hp:tbl']);

                    if (!hasTbl) {
                        const pContentHtml = this.extractRunsFromPara(p, imageMap, false, styleCharClass);
                        html += `<div class="${paraClass}" data-hwpx="${pPath}">${pContentHtml}</div>`;
                    } else {
                        const textHtml = this.extractRunsFromPara(p, imageMap, true, styleCharClass);
                        if (textHtml.trim()) {
                            html += `<div class="${paraClass}" data-hwpx="${pPath}">${textHtml}</div>`;
                        }
                        for (let rri = 0; rri < runs.length; rri++) {
                            if (runs[rri]['hp:tbl']) {
                                html += this.extractHtml({ 'hp:tbl': runs[rri]['hp:tbl'] }, imageMap, `${pPath}/hp:run[${rri}]`);
                            }
                        }
                    }
                }
            }
            // 그 외 노드 재귀 탐색
            else {
                for (const key in obj) {
                    if (!key.startsWith('@_')) {
                        html += this.extractHtml(obj[key], imageMap, xmlPath);
                    }
                }
            }
        }

        return html;
    }

    /**
     * 문단 객체 내부(hp:run)의 텍스트, 이미지, 컨트롤(하이퍼링크/자동번호) 등을 HTML로 변환합니다.
     */
    private static extractRunsFromPara(pObj: any, imageMap: Record<string, string>, skipTables: boolean = false, styleCharClass: string = ''): string {
        let html = '';
        if (!pObj) return html;

        const runs = Array.isArray(pObj['hp:run']) ? pObj['hp:run'] : (pObj['hp:run'] ? [pObj['hp:run']] : []);

        // 하이퍼링크 필드 상태 추적
        let activeHyperlink: string | null = null;
        let activeFieldId: string | null = null;

        for (const run of runs) {
            const charId = run['@_charPrIDRef'] !== undefined ? run['@_charPrIDRef'] : '';
            const charClass = charId !== '' ? `char-${charId}` : styleCharClass;

            const ctrl = run['hp:ctrl'];

            // 텍스트 조각 수집
            const ts = Array.isArray(run['hp:t']) ? run['hp:t'] : (run['hp:t'] ? [run['hp:t']] : []);
            const textFragments: string[] = [];
            for (const t of ts) {
                let frag = '';
                if (typeof t === 'string' || typeof t === 'number') {
                    frag = String(t);
                } else if (typeof t === 'object' && t !== null) {
                    if (t['#text'] !== undefined) frag += t['#text'];
                    if (t['hp:lineBreak'] !== undefined) frag += '\n';
                }
                textFragments.push(frag);
            }

            const escapeText = (s: string) => {
                let result = s
                    .replace(/&/g, "&amp;")
                    .replace(/</g, "&lt;")
                    .replace(/>/g, "&gt;")
                    .replace(/ {2,}/g, (m: string) => '&nbsp;'.repeat(m.length))
                    .replace(/\t/g, '<span style="display:inline-block; min-width:2em;">\u00A0</span>')
                    .replace(/\n/g, "<br>");
                // 마크펜 (형광펜) 마커 → <mark> 태그
                result = result.replace(/\uE000([^\uE001]*)\uE001/g, '<mark style="background-color:$1">');
                result = result.replace(/\uE002/g, '</mark>');
                // 페이지 번호 마커
                result = result.replace(/\uE003/g, '<span class="hwpx-pagenum"></span>');
                return result;
            };

            const outputText = (text: string) => {
                const escaped = escapeText(text);
                if (escaped) html += `<span class="${charClass}">${escaped}</span>`;
            };

            // hp:ctrl 처리 (하이퍼링크, 자동번호 등)
            const processCtrl = () => {
                if (!ctrl) return;
                const fieldBegin = ctrl['hp:fieldBegin'];
                if (fieldBegin) {
                    const fieldType = fieldBegin['@_type'];
                    if (fieldType === 'HYPERLINK') {
                        activeFieldId = fieldBegin['@_id'];
                        const params = fieldBegin['hp:parameters'];
                        if (params) {
                            const stringParams = Array.isArray(params['hp:stringParam']) ? params['hp:stringParam'] : (params['hp:stringParam'] ? [params['hp:stringParam']] : []);
                            for (const sp of stringParams) {
                                if (sp['@_name'] === 'Path') {
                                    activeHyperlink = sp['#text'] || '';
                                    break;
                                }
                            }
                        }
                        if (activeHyperlink) {
                            html += `<a href="${activeHyperlink}" style="text-decoration: none; color: inherit;" target="_blank">`;
                        }
                    }
                }
                const fieldEnd = ctrl['hp:fieldEnd'];
                if (fieldEnd) {
                    const beginRef = fieldEnd['@_beginIDRef'];
                    if (activeHyperlink && activeFieldId === beginRef) {
                        html += `</a>`;
                        activeHyperlink = null;
                        activeFieldId = null;
                    }
                }
                if (ctrl['hp:autoNum']) {
                    const num = String(ctrl['hp:autoNum']['@_num'] || '');
                    if (num) html += `<span class="${charClass}">${escapeText(num)}</span>`;
                }
                // 각주 (footNote)
                if (ctrl['hp:footNote']) {
                    const fn = ctrl['hp:footNote'];
                    const fnSubList = fn['hp:subList'];
                    let fnNum = '';
                    // 각주 번호: autoNum에서 추출 또는 instId 기반
                    if (fnSubList) {
                        const fnParas = Array.isArray(fnSubList['hp:p']) ? fnSubList['hp:p'] : (fnSubList['hp:p'] ? [fnSubList['hp:p']] : []);
                        for (const fp of fnParas) {
                            const fRuns = Array.isArray(fp['hp:run']) ? fp['hp:run'] : (fp['hp:run'] ? [fp['hp:run']] : []);
                            for (const fr of fRuns) {
                                if (fr['hp:ctrl']?.['hp:autoNum']) {
                                    fnNum = String(fr['hp:ctrl']['hp:autoNum']['@_num'] || '');
                                    break;
                                }
                            }
                            if (fnNum) break;
                        }
                    }
                    if (!fnNum) fnNum = String(fn['@_instId'] || '?');
                    html += `<sup class="hwpx-fn-ref" style="color: blue; cursor: default;">${escapeText(fnNum)}</sup>`;
                    // 각주 내용을 숨긴 div로 추가 (페이지네이션 JS에서 페이지 하단에 배치)
                    if (fnSubList) {
                        const fnContent = HwpxParser.extractHtml(fnSubList, imageMap);
                        html += `<div class="hwpx-footnote-content" style="display:none;" data-fn-num="${escapeText(fnNum)}">${fnContent}</div>`;
                    }
                }
                // 미주 (endNote)
                if (ctrl['hp:endNote']) {
                    const en = ctrl['hp:endNote'];
                    const enSubList = en['hp:subList'];
                    let enNum = '';
                    if (enSubList) {
                        const enParas = Array.isArray(enSubList['hp:p']) ? enSubList['hp:p'] : (enSubList['hp:p'] ? [enSubList['hp:p']] : []);
                        for (const ep of enParas) {
                            const eRuns = Array.isArray(ep['hp:run']) ? ep['hp:run'] : (ep['hp:run'] ? [ep['hp:run']] : []);
                            for (const er of eRuns) {
                                if (er['hp:ctrl']?.['hp:autoNum']) {
                                    enNum = String(er['hp:ctrl']['hp:autoNum']['@_num'] || '');
                                    break;
                                }
                            }
                            if (enNum) break;
                        }
                    }
                    if (!enNum) enNum = String(en['@_instId'] || '?');
                    html += `<sup class="hwpx-en-ref" style="color: blue; cursor: default;">${escapeText(enNum)}</sup>`;
                    if (enSubList) {
                        const enContent = HwpxParser.extractHtml(enSubList, imageMap);
                        html += `<div class="hwpx-endnote-content" style="display:none;" data-en-num="${escapeText(enNum)}">${enContent}</div>`;
                    }
                }
                // 다단 설정 (colPr)
                if (ctrl['hp:colPr']) {
                    const colPr = ctrl['hp:colPr'];
                    const colCount = parseInt(colPr['@_colCount'] || '1');
                    if (colCount > 1) {
                        const gap = parseInt(colPr['@_sameGap'] || '0') / 283.465;
                        const colLine = colPr['hp:colLine'];
                        let ruleCss = '';
                        if (colLine) {
                            const lineType = colLine['@_type'] || 'NONE';
                            const lineColor = colLine['@_color'] || '#000000';
                            if (lineType !== 'NONE') {
                                ruleCss = `column-rule: 1px solid ${lineColor};`;
                            }
                        }
                        html += `<div style="column-count: ${colCount}; column-gap: ${gap.toFixed(1)}mm; ${ruleCss}">`;
                    }
                }
                // 단 구분 (columnBreak) - hp:t 내에서 전처리하기 어려우므로 ctrl 레벨에서 처리
                if (ctrl['hp:columnBreak'] !== undefined) {
                    html += `<div style="break-before: column;"></div>`;
                }
            };

            // HWPX에서 hp:t와 hp:ctrl의 문서 순서를 복원하여 출력
            const hasCtrl = ctrl !== undefined && ctrl !== null;
            if (textFragments.length >= 2 && hasCtrl) {
                // 다중 텍스트 + ctrl: t[0] → ctrl → t[1..] (캡션 autoNum 등)
                outputText(textFragments[0]);
                processCtrl();
                outputText(textFragments.slice(1).join(''));
            } else if (hasCtrl && textFragments.length > 0) {
                const hasField = ctrl['hp:fieldBegin'] || ctrl['hp:fieldEnd'];
                if (hasField) {
                    // 필드 관련 ctrl: HWPX에서 텍스트가 필드 앞에 위치
                    outputText(textFragments.join(''));
                    processCtrl();
                } else {
                    // autoNum 등: ctrl 먼저
                    processCtrl();
                    outputText(textFragments.join(''));
                }
            } else if (hasCtrl) {
                processCtrl();
            } else {
                outputText(textFragments.join(''));
            }

            // 표 (run 레벨)
            if (run['hp:tbl'] && !skipTables) {
                html += this.extractHtml({ 'hp:tbl': run['hp:tbl'] }, imageMap);
            }

            // 그림(hp:pic) 처리
            if (run['hp:pic']) {
                const pics = Array.isArray(run['hp:pic']) ? run['hp:pic'] : [run['hp:pic']];
                for (const pic of pics) {
                    const imgNode = this.findDeepNode(pic, 'hc:img');
                    if (imgNode) {
                        const binItemIdRef = imgNode['@_binaryItemIDRef'] || imgNode['@_binItem'];
                        if (binItemIdRef && imageMap[binItemIdRef]) {
                            let imgStyle = '';

                            // 이미지 크기 (curSz) — max-width로 컨테이너 내 제한
                            const curSz = pic['hp:curSz'];
                            if (curSz) {
                                const imgW = parseInt(curSz['@_width'] || '0') / 283.465;
                                const imgH = parseInt(curSz['@_height'] || '0') / 283.465;
                                if (imgW > 0) imgStyle += `width: ${imgW.toFixed(1)}mm; max-width: 100%; `;
                                if (imgH > 0) imgStyle += `height: auto; `;
                            } else {
                                imgStyle += `max-width: 100%; height: auto; `;
                            }

                            // 회전 + 플립 → transform
                            let transforms: string[] = [];
                            const rotationInfo = pic['hp:rotationInfo'];
                            if (rotationInfo) {
                                const angle = parseInt(rotationInfo['@_angle'] || '0');
                                if (angle !== 0) transforms.push(`rotate(${angle}deg)`);
                            }
                            const flip = pic['hp:flip'];
                            if (flip) {
                                if (flip['@_horizontal'] === '1') transforms.push('scaleX(-1)');
                                if (flip['@_vertical'] === '1') transforms.push('scaleY(-1)');
                            }
                            if (transforms.length > 0) {
                                imgStyle += `transform: ${transforms.join(' ')}; `;
                            }

                            // 이미지 크롭 (imgClip + imgDim)
                            const imgClip = pic['hp:imgClip'];
                            const imgDim = pic['hp:imgDim'];
                            if (imgClip && imgDim) {
                                const dimW = parseInt(imgDim['@_dimwidth'] || '0');
                                const dimH = parseInt(imgDim['@_dimheight'] || '0');
                                const clipL = parseInt(imgClip['@_left'] || '0');
                                const clipR = parseInt(imgClip['@_right'] || '0');
                                const clipT = parseInt(imgClip['@_top'] || '0');
                                const clipB = parseInt(imgClip['@_bottom'] || '0');
                                if (dimW > 0 && dimH > 0 && (clipL > 0 || clipT > 0 || clipR < dimW || clipB < dimH)) {
                                    const pctL = (clipL / dimW) * 100;
                                    const pctT = (clipT / dimH) * 100;
                                    const pctR = (clipR / dimW) * 100;
                                    const pctB = (clipB / dimH) * 100;
                                    imgStyle += `clip-path: inset(${pctT.toFixed(1)}% ${(100 - pctR).toFixed(1)}% ${(100 - pctB).toFixed(1)}% ${pctL.toFixed(1)}%); `;
                                }
                            }

                            // 이미지 여백 (outMargin → margin, inMargin → padding)
                            const outMargin = pic['hp:outMargin'];
                            if (outMargin) {
                                const omT = (parseInt(outMargin['@_top'] || '0') / 283.465).toFixed(1);
                                const omB = (parseInt(outMargin['@_bottom'] || '0') / 283.465).toFixed(1);
                                const omL = (parseInt(outMargin['@_left'] || '0') / 283.465).toFixed(1);
                                const omR = (parseInt(outMargin['@_right'] || '0') / 283.465).toFixed(1);
                                if (parseFloat(omT) || parseFloat(omB) || parseFloat(omL) || parseFloat(omR)) {
                                    imgStyle += `margin: ${omT}mm ${omR}mm ${omB}mm ${omL}mm; `;
                                }
                            }

                            // 이미지 배치 (pos)
                            const pos = pic['hp:pos'];
                            let posStyle = 'display: inline-block; vertical-align: middle; ';
                            if (pos) {
                                const treatAsChar = pos['@_treatAsChar'];
                                if (treatAsChar === '0') {
                                    const horzAlign = (pos['@_horzAlign'] || 'LEFT').toUpperCase();
                                    if (horzAlign === 'CENTER') {
                                        posStyle = 'display: block; margin-left: auto; margin-right: auto; ';
                                    } else if (horzAlign === 'RIGHT') {
                                        posStyle = 'display: block; margin-left: auto; margin-right: 0; ';
                                    } else {
                                        posStyle = 'display: block; ';
                                    }
                                }
                            }

                            // 캡션 (caption)
                            const caption = pic['hp:caption'];
                            if (caption) {
                                const captionSide = (caption['@_side'] || 'BOTTOM').toUpperCase();
                                const captionGap = (parseInt(caption['@_gap'] || '850') / 283.465).toFixed(1);
                                let captionHtml = '';
                                const captionSubList = caption['hp:subList'];
                                if (captionSubList) {
                                    captionHtml = this.extractHtml(captionSubList, imageMap);
                                }

                                html += `<figure style="${posStyle} text-align: center;">`;
                                if (captionSide === 'TOP') {
                                    html += `<figcaption style="margin-bottom: ${captionGap}mm; ">${captionHtml}</figcaption>`;
                                }
                                html += `<img src="${imageMap[binItemIdRef]}" style="${imgStyle}"/>`;
                                if (captionSide !== 'TOP') {
                                    html += `<figcaption style="margin-top: ${captionGap}mm; ">${captionHtml}</figcaption>`;
                                }
                                html += `</figure>`;
                            } else {
                                html += `<img src="${imageMap[binItemIdRef]}" style="${posStyle} ${imgStyle}"/>`;
                            }
                        } else {
                            html += `<span style="border:1px dashed #f00; padding:2px;">[이미지 누락: ${binItemIdRef}]</span>`;
                        }
                    } else {
                        html += `<span style="border:1px dashed #f00; padding:2px;">[그림 객체]</span>`;
                    }
                }
            }

            // 도형 객체들 (line, rect, ellipse, polygon, arc, curve, connectLine)
            const shapeTypes = ['hp:line', 'hp:rect', 'hp:ellipse', 'hp:polygon', 'hp:arc', 'hp:curve', 'hp:connectLine'];
            for (const shapeType of shapeTypes) {
                if (run[shapeType]) {
                    const shapes = Array.isArray(run[shapeType]) ? run[shapeType] : [run[shapeType]];
                    for (const shape of shapes) {
                        html += this.renderDrawingObject(shape, shapeType, imageMap);
                    }
                }
            }

            // 그룹 도형 (container)
            if (run['hp:container']) {
                const containers = Array.isArray(run['hp:container']) ? run['hp:container'] : [run['hp:container']];
                for (const container of containers) {
                    html += this.renderContainer(container, imageMap);
                }
            }

            // 수식 (equation)
            if (run['hp:equation']) {
                const eqs = Array.isArray(run['hp:equation']) ? run['hp:equation'] : [run['hp:equation']];
                for (const eq of eqs) {
                    const script = eq['hp:script'] || eq['#text'] || '';
                    const sz = eq['hp:sz'];
                    let eqStyle = 'display: inline-block; vertical-align: middle; padding: 4px; font-family: "Cambria Math", "Latin Modern Math", serif; ';
                    if (sz) {
                        const w = parseInt(sz['@_width'] || '0') / 283.465;
                        const h = parseInt(sz['@_height'] || '0') / 283.465;
                        if (w > 0) eqStyle += `width: ${w.toFixed(1)}mm; max-width: 100%; `;
                        if (h > 0) eqStyle += `min-height: ${h.toFixed(1)}mm; `;
                    }
                    html += `<code class="hwpx-equation" style="${eqStyle}">${this.escapeHtml(String(script))}</code>`;
                }
            }

            // OLE 객체
            if (run['hp:ole']) {
                const oles = Array.isArray(run['hp:ole']) ? run['hp:ole'] : [run['hp:ole']];
                for (const ole of oles) {
                    const binRef = ole['@_binaryItemIDRef'];
                    const sz = ole['hp:sz'];
                    let oleStyle = 'display: inline-block; vertical-align: middle; ';
                    if (sz) {
                        const w = parseInt(sz['@_width'] || '0') / 283.465;
                        const h = parseInt(sz['@_height'] || '0') / 283.465;
                        if (w > 0) oleStyle += `width: ${w.toFixed(1)}mm; max-width: 100%; `;
                        if (h > 0) oleStyle += `height: ${h.toFixed(1)}mm; `;
                    }
                    if (binRef && imageMap[binRef]) {
                        html += `<img src="${imageMap[binRef]}" style="${oleStyle}" alt="OLE 개체"/>`;
                    } else {
                        html += `<span style="border: 1px dashed #999; padding: 4px; ${oleStyle}">[OLE 개체]</span>`;
                    }
                }
            }

            // 텍스트 아트 (textart)
            if (run['hp:textart']) {
                const textarts = Array.isArray(run['hp:textart']) ? run['hp:textart'] : [run['hp:textart']];
                for (const ta of textarts) {
                    const sz = ta['hp:sz'];
                    let taW = 50, taH = 20;
                    if (sz) {
                        taW = parseInt(sz['@_width'] || '14174') / 283.465;
                        taH = parseInt(sz['@_height'] || '5670') / 283.465;
                    }
                    const taPr = ta['hp:textartPr'];
                    const text = ta['#text'] || taPr?.['#text'] || 'TextArt';
                    let fontStyle = '';
                    if (taPr) {
                        if (taPr['@_fontBold'] === '1') fontStyle += 'font-weight: bold; ';
                        if (taPr['@_fontItalic'] === '1') fontStyle += 'font-style: italic; ';
                    }
                    html += `<svg width="${taW}mm" height="${taH}mm" viewBox="0 0 ${taW} ${taH}" style="display: inline-block; vertical-align: middle; max-width: 100%;">`;
                    html += `<text x="50%" y="60%" text-anchor="middle" dominant-baseline="middle" style="font-size: ${(taH * 0.7).toFixed(0)}px; ${fontStyle}">${this.escapeHtml(String(text))}</text>`;
                    html += `</svg>`;
                }
            }

            // 폼 컨트롤 (form controls)
            if (run['hp:checkBtn']) {
                const cb = run['hp:checkBtn'];
                const checked = cb['@_value'] === '1' || cb['@_triState'] === '1';
                html += `<input type="checkbox" disabled ${checked ? 'checked' : ''} style="vertical-align: middle;"/>`;
            }
            if (run['hp:radioBtn']) {
                const rb = run['hp:radioBtn'];
                const checked = rb['@_value'] === '1';
                html += `<input type="radio" disabled ${checked ? 'checked' : ''} style="vertical-align: middle;"/>`;
            }
            if (run['hp:edit']) {
                const ed = run['hp:edit'];
                const defaultVal = ed['@_defaultValue'] || ed['#text'] || '';
                html += `<span style="border: 1px solid #999; padding: 1px 4px; display: inline-block; min-width: 3em;">${this.escapeHtml(String(defaultVal))}</span>`;
            }
            if (run['hp:comboBox']) {
                const combo = run['hp:comboBox'];
                const selected = combo['@_selectedValue'] || '';
                const items = Array.isArray(combo['hp:listItem']) ? combo['hp:listItem'] : (combo['hp:listItem'] ? [combo['hp:listItem']] : []);
                const displayText = items.find((i: any) => i['@_value'] === selected)?.['@_displayText'] || selected;
                html += `<span style="border: 1px solid #999; padding: 1px 4px; display: inline-block;">${this.escapeHtml(String(displayText))} ▾</span>`;
            }
            if (run['hp:btn']) {
                const btn = run['hp:btn'];
                const caption = btn['@_caption'] || 'Button';
                html += `<span style="border: 1px solid #999; padding: 2px 8px; background: #f0f0f0; display: inline-block;">${this.escapeHtml(String(caption))}</span>`;
            }

            // 덧말 (dutmal / ruby)
            if (run['hp:dutmal']) {
                const dm = run['hp:dutmal'];
                const mainText = dm['hp:mainText'] || '';
                const subText = dm['hp:subText'] || '';
                html += `<ruby><span class="${charClass}">${this.escapeHtml(String(mainText))}</span><rp>(</rp><rt>${this.escapeHtml(String(subText))}</rt><rp>)</rp></ruby>`;
            }

            // 글자 겹침 (compose)
            if (run['hp:compose']) {
                const comp = run['hp:compose'];
                const composeText = comp['@_composeText'] || '';
                html += `<span class="${charClass}" style="display: inline-block; position: relative;">${this.escapeHtml(String(composeText))}</span>`;
            }

            // 숨은 설명 (hiddenComment)
            if (run['hp:hiddenComment']) {
                const hc = run['hp:hiddenComment'];
                const subList = hc['hp:subList'];
                let commentText = '';
                if (subList) {
                    commentText = this.extractHtml(subList, imageMap).replace(/"/g, '&quot;');
                }
                html += `<span class="hwpx-comment" title="${commentText}" style="cursor: help; border-bottom: 1px dotted #999;">*</span>`;
            }
        }

        return html;
    }

    /**
     * 도형 객체를 SVG로 렌더링합니다.
     */
    private static renderDrawingObject(shape: any, shapeType: string, imageMap: Record<string, string>): string {
        const sz = shape['hp:sz'];
        let w = 50, h = 50;
        if (sz) {
            w = parseInt(sz['@_width'] || '14174') / 283.465;
            h = parseInt(sz['@_height'] || '14174') / 283.465;
        }

        // 위치/배치
        const pos = shape['hp:pos'];
        let wrapStyle = 'display: inline-block; vertical-align: middle; ';
        if (pos && pos['@_treatAsChar'] === '0') {
            const hAlign = (pos['@_horzAlign'] || 'LEFT').toUpperCase();
            if (hAlign === 'CENTER') wrapStyle = 'display: block; margin: 0 auto; ';
            else if (hAlign === 'RIGHT') wrapStyle = 'display: block; margin-left: auto; margin-right: 0; ';
            else wrapStyle = 'display: block; ';
        }

        // 변환 (회전 + 플립)
        let transforms: string[] = [];
        const rotInfo = shape['hp:rotationInfo'];
        if (rotInfo) {
            const angle = parseInt(rotInfo['@_angle'] || '0');
            if (angle !== 0) transforms.push(`rotate(${angle}deg)`);
        }
        const flip = shape['hp:flip'];
        if (flip) {
            if (flip['@_horizontal'] === '1') transforms.push('scaleX(-1)');
            if (flip['@_vertical'] === '1') transforms.push('scaleY(-1)');
        }
        const transformCss = transforms.length > 0 ? `transform: ${transforms.join(' ')};` : '';

        // 선 스타일
        const lineShape = shape['hp:lineShape'];
        let stroke = '#000000', strokeWidth = '0.5', strokeDash = '';
        if (lineShape) {
            stroke = lineShape['@_color'] || '#000000';
            const lw = parseFloat(lineShape['@_width'] || '0.12 mm');
            strokeWidth = String(lw > 0 ? lw : 0.5);
            const lineType = lineShape['@_type'] || 'SOLID';
            if (lineType === 'DOTTED' || lineType === 'DASH_DOT') strokeDash = 'stroke-dasharray="2,2"';
            else if (lineType === 'DASHED') strokeDash = 'stroke-dasharray="6,3"';
            else if (lineType === 'NONE') { stroke = 'none'; strokeWidth = '0'; }
        }

        // 채움 (fillBrush)
        let fill = 'none';
        const fillBrush = shape['hp:fillBrush'] || shape['hc:fillBrush'];
        if (fillBrush) {
            if (fillBrush['hc:winBrush']) {
                fill = fillBrush['hc:winBrush']['@_faceColor'] || 'none';
            }
        }

        let svgContent = '';
        const vbW = w, vbH = h;

        switch (shapeType) {
            case 'hp:line': {
                const startPt = shape['hp:startPt'];
                const endPt = shape['hp:endPt'];
                const x1 = startPt ? parseInt(startPt['@_x'] || '0') / 283.465 : 0;
                const y1 = startPt ? parseInt(startPt['@_y'] || '0') / 283.465 : 0;
                const x2 = endPt ? parseInt(endPt['@_x'] || '0') / 283.465 : w;
                const y2 = endPt ? parseInt(endPt['@_y'] || '0') / 283.465 : h;
                svgContent = `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${stroke}" stroke-width="${strokeWidth}" ${strokeDash}/>`;
                break;
            }
            case 'hp:rect': {
                svgContent = `<rect x="0" y="0" width="${vbW}" height="${vbH}" fill="${fill}" stroke="${stroke}" stroke-width="${strokeWidth}" ${strokeDash}/>`;
                break;
            }
            case 'hp:ellipse': {
                svgContent = `<ellipse cx="${vbW / 2}" cy="${vbH / 2}" rx="${vbW / 2}" ry="${vbH / 2}" fill="${fill}" stroke="${stroke}" stroke-width="${strokeWidth}" ${strokeDash}/>`;
                break;
            }
            case 'hp:polygon': {
                const pts = Array.isArray(shape['hp:pt']) ? shape['hp:pt'] : (shape['hp:pt'] ? [shape['hp:pt']] : []);
                const points = pts.map((p: any) => `${parseInt(p['@_x'] || '0') / 283.465},${parseInt(p['@_y'] || '0') / 283.465}`).join(' ');
                svgContent = `<polygon points="${points}" fill="${fill}" stroke="${stroke}" stroke-width="${strokeWidth}" ${strokeDash}/>`;
                break;
            }
            case 'hp:arc': {
                svgContent = `<ellipse cx="${vbW / 2}" cy="${vbH / 2}" rx="${vbW / 2}" ry="${vbH / 2}" fill="${fill}" stroke="${stroke}" stroke-width="${strokeWidth}" ${strokeDash}/>`;
                break;
            }
            case 'hp:curve': {
                const segs = Array.isArray(shape['hp:seg']) ? shape['hp:seg'] : (shape['hp:seg'] ? [shape['hp:seg']] : []);
                if (segs.length > 0) {
                    let d = `M 0 0`;
                    for (const seg of segs) {
                        const type = seg['@_type'] || 'CURVE';
                        const x1 = parseInt(seg['@_x1'] || '0') / 283.465;
                        const y1 = parseInt(seg['@_y1'] || '0') / 283.465;
                        const x2 = parseInt(seg['@_x2'] || '0') / 283.465;
                        const y2 = parseInt(seg['@_y2'] || '0') / 283.465;
                        if (type === 'LINE') d += ` L ${x2} ${y2}`;
                        else d += ` C ${x1} ${y1} ${x2} ${y2} ${x2} ${y2}`;
                    }
                    svgContent = `<path d="${d}" fill="${fill}" stroke="${stroke}" stroke-width="${strokeWidth}" ${strokeDash}/>`;
                }
                break;
            }
            case 'hp:connectLine': {
                const sp = shape['hp:startPt'];
                const ep = shape['hp:endPt'];
                const sx = sp ? parseInt(sp['@_x'] || '0') / 283.465 : 0;
                const sy = sp ? parseInt(sp['@_y'] || '0') / 283.465 : 0;
                const ex = ep ? parseInt(ep['@_x'] || '0') / 283.465 : w;
                const ey = ep ? parseInt(ep['@_y'] || '0') / 283.465 : h;
                svgContent = `<line x1="${sx}" y1="${sy}" x2="${ex}" y2="${ey}" stroke="${stroke}" stroke-width="${strokeWidth}" ${strokeDash}/>`;
                break;
            }
        }

        // 텍스트 박스 (도형 내부 텍스트)
        let textboxHtml = '';
        const textbox = shape['hp:textbox'];
        if (textbox) {
            const subList = textbox['hp:subList'];
            if (subList) {
                textboxHtml = this.extractHtml(subList, imageMap);
            }
        }

        let result = `<div style="${wrapStyle} width: ${w.toFixed(1)}mm; max-width: 100%; ${transformCss}">`;
        result += `<svg width="100%" height="${h.toFixed(1)}mm" viewBox="0 0 ${vbW.toFixed(1)} ${vbH.toFixed(1)}" preserveAspectRatio="none" xmlns="http://www.w3.org/2000/svg">`;
        result += svgContent;
        if (textboxHtml) {
            result += `<foreignObject x="0" y="0" width="${vbW.toFixed(1)}" height="${vbH.toFixed(1)}"><div xmlns="http://www.w3.org/1999/xhtml" style="width:100%; height:100%; display:flex; align-items:center; justify-content:center; overflow:hidden; padding: 2px; box-sizing: border-box;">${textboxHtml}</div></foreignObject>`;
        }
        result += `</svg></div>`;
        return result;
    }

    /**
     * 그룹 도형 (container)을 렌더링합니다.
     */
    private static renderContainer(container: any, imageMap: Record<string, string>): string {
        const sz = container['hp:sz'];
        let w = 50, h = 50;
        if (sz) {
            w = parseInt(sz['@_width'] || '14174') / 283.465;
            h = parseInt(sz['@_height'] || '14174') / 283.465;
        }

        let html = `<div style="display: inline-block; vertical-align: middle; position: relative; width: ${w.toFixed(1)}mm; height: ${h.toFixed(1)}mm; max-width: 100%;">`;

        // 자식 도형 렌더링
        const shapeTypes = ['hp:line', 'hp:rect', 'hp:ellipse', 'hp:polygon', 'hp:arc', 'hp:curve', 'hp:connectLine'];
        for (const st of shapeTypes) {
            if (container[st]) {
                const shapes = Array.isArray(container[st]) ? container[st] : [container[st]];
                for (const shape of shapes) {
                    html += this.renderDrawingObject(shape, st, imageMap);
                }
            }
        }
        // 중첩 컨테이너
        if (container['hp:container']) {
            const subs = Array.isArray(container['hp:container']) ? container['hp:container'] : [container['hp:container']];
            for (const sub of subs) {
                html += this.renderContainer(sub, imageMap);
            }
        }
        // 그룹 내 이미지
        if (container['hp:pic']) {
            const pics = Array.isArray(container['hp:pic']) ? container['hp:pic'] : [container['hp:pic']];
            for (const pic of pics) {
                const imgNode = this.findDeepNode(pic, 'hc:img');
                if (imgNode) {
                    const binRef = imgNode['@_binaryItemIDRef'] || imgNode['@_binItem'];
                    if (binRef && imageMap[binRef]) {
                        html += `<img src="${imageMap[binRef]}" style="max-width: 100%; height: auto;"/>`;
                    }
                }
            }
        }

        html += `</div>`;
        return html;
    }

    /**
     * HTML 특수문자를 이스케이프합니다.
     */
    private static escapeHtml(s: string): string {
        return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    /**
     * 특정 노드를 재귀적으로 찾아 반환합니다.
     */
    private static findDeepNode(obj: any, targetKey: string): any {
        if (!obj || typeof obj !== 'object') return null;
        if (targetKey in obj) return obj[targetKey];

        for (const key in obj) {
            if (!key.startsWith('@_')) {
                const result = this.findDeepNode(obj[key], targetKey);
                if (result) return result;
            }
        }
        return null;
    }

    // ===== Border 헬퍼 메서드 =====

    private static borderTypeToCss(type: string): string {
        switch (type) {
            case 'SOLID': return 'solid';
            case 'DOTTED': case 'DASH_DOT': case 'DASH_DOT_DOT': return 'dotted';
            case 'DASHED': return 'dashed';
            case 'DOUBLE': return 'double';
            case 'NONE': return 'none';
            default: return 'solid';
        }
    }

    private static parseBorderWidth(widthStr: string): string {
        const match = widthStr.match(/([\d.]+)\s*mm/);
        if (match) return `${match[1]}mm`;
        return '0.1mm';
    }

    private static borderInfoToCss(info: BorderInfo): string {
        if (info.type === 'NONE') return 'none';
        return `${this.parseBorderWidth(info.width)} ${this.borderTypeToCss(info.type)} ${info.color}`;
    }
}
