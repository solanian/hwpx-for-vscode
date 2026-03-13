import * as http from 'http';
import * as crypto from 'crypto';
import * as vscode from 'vscode';
import JSZip from 'jszip';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

interface XpathSegment {
    type: 'child' | 'descendant';
    part: { name: string; index: number };
}

interface OpenDocument {
    uri: vscode.Uri;
    zip: JSZip;
    dirty: boolean;
}

export class HwpxApiServer {
    private server: http.Server | null = null;
    private port: number = 0;
    private token: string = '';
    private documents: Map<string, OpenDocument> = new Map();
    private reloadListeners: Map<string, () => void> = new Map();

    private parser = new XMLParser({
        ignoreAttributes: false,
        attributeNamePrefix: "@_",
        textNodeName: "#text",
        trimValues: false,
        preserveOrder: true,
    });

    private builder = new XMLBuilder({
        ignoreAttributes: false,
        attributeNamePrefix: "@_",
        textNodeName: "#text",
        preserveOrder: true,
        format: false,
        suppressEmptyNode: true,
    });

    async start(): Promise<number> {
        this.token = crypto.randomBytes(32).toString('hex');
        return new Promise((resolve, reject) => {
            this.server = http.createServer((req, res) => this.handleRequest(req, res));
            this.server.listen(0, '127.0.0.1', () => {
                const addr = this.server!.address();
                if (addr && typeof addr === 'object') {
                    this.port = addr.port;
                    console.log(`HWPX API server listening on http://127.0.0.1:${this.port}`);
                    resolve(this.port);
                } else {
                    reject(new Error('Failed to get server address'));
                }
            });
            this.server.on('error', reject);
        });
    }

    getToken(): string {
        return this.token;
    }

    stop() {
        if (this.server) {
            this.server.close();
            this.server = null;
        }
    }

    getPort(): number {
        return this.port;
    }

    registerDocument(uri: vscode.Uri, zip: JSZip) {
        const key = uri.fsPath;
        this.documents.set(key, { uri, zip, dirty: false });
    }

    unregisterDocument(uri: vscode.Uri) {
        this.documents.delete(uri.fsPath);
        this.reloadListeners.delete(uri.fsPath);
    }

    onReload(uri: vscode.Uri, listener: () => void) {
        this.reloadListeners.set(uri.fsPath, listener);
    }

    private triggerReload(docPath: string) {
        const listener = this.reloadListeners.get(docPath);
        if (listener) listener();
    }

    private async handleRequest(req: http.IncomingMessage, res: http.ServerResponse) {
        const origin = req.headers.origin || '';
        const allowedOrigins = ['vscode-webview://', 'http://127.0.0.1', 'http://localhost'];
        const isAllowedOrigin = !origin || allowedOrigins.some(o => origin.startsWith(o));

        // CORS headers
        res.setHeader('Access-Control-Allow-Origin', isAllowedOrigin ? (origin || 'http://127.0.0.1') : 'null');
        res.setHeader('Access-Control-Allow-Methods', 'GET, PUT, POST, OPTIONS');
        res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
        res.setHeader('Vary', 'Origin');

        if (req.method === 'OPTIONS') {
            res.writeHead(200, { 'Content-Length': '0', 'Connection': 'close' });
            res.end();
            return;
        }

        const url = new URL(req.url || '/', `http://127.0.0.1:${this.port}`);
        const pathname = url.pathname;

        // /api/help는 토큰 없이 접근 가능 (문서 조회용)
        if (pathname === '/api/help' && req.method === 'GET') {
            try {
                await this.handleHelp(res);
            } catch (err: any) {
                this.sendJson(res, 500, { error: err.message });
            }
            return;
        }

        // 토큰 인증: Authorization 헤더 또는 ?token= 쿼리 파라미터
        const authHeader = req.headers.authorization || '';
        const queryToken = url.searchParams.get('token') || '';
        const providedToken = authHeader.replace(/^Bearer\s+/i, '') || queryToken;

        if (providedToken !== this.token) {
            this.sendJson(res, 401, { error: 'Unauthorized. Provide token via Authorization: Bearer <token> header or ?token= query parameter.' });
            return;
        }

        try {
            if (pathname === '/api/documents' && req.method === 'GET') {
                await this.handleListDocuments(res);
            } else if (pathname === '/api/files' && req.method === 'GET') {
                await this.handleListFiles(url, res);
            } else if (pathname === '/api/element' && req.method === 'GET') {
                await this.handleGetElement(url, res);
            } else if (pathname === '/api/element' && req.method === 'PUT') {
                await this.handlePutElement(url, req, res);
            } else if (pathname === '/api/save' && req.method === 'POST') {
                await this.handleSave(url, res);
            } else if (pathname === '/api/reload' && req.method === 'POST') {
                await this.handleReload(url, res);
            } else if (pathname === '/api/xml' && req.method === 'GET') {
                await this.handleGetXml(url, res);
            } else {
                this.sendJson(res, 404, { error: 'Not found. GET /api/help for usage.' });
            }
        } catch (err: any) {
            this.sendJson(res, 500, { error: err.message });
        }
    }

    /** GET /api/help — API 사용법 문서 (토큰 불필요) */
    private async handleHelp(res: http.ServerResponse) {
        const base = `http://127.0.0.1:${this.port}`;
        const docs = Array.from(this.documents.keys());

        const helpText = `# HWPX API Reference

Base URL: ${base}
Token: Required for all endpoints except /api/help.
  - Header: Authorization: Bearer <token>
  - Query:  ?token=<token>

## Endpoints

### GET /api/help
API 사용법 (이 문서). 토큰 불필요.

### GET /api/documents
열려있는 HWPX 문서 목록 조회.
Response: { "documents": [{ "path": "<filePath>", "dirty": <bool> }] }

### GET /api/files?doc=<path>
HWPX(ZIP) 내부의 XML 파일 목록 조회.
- doc: 문서 경로 (문서가 1개면 생략 가능)
Response: { "files": ["Contents/section0.xml", ...] }

### GET /api/xml?doc=<path>&file=<xmlPath>
XML 파일 원본 내용을 그대로 반환.
- file: ZIP 내부 경로 (예: Contents/section0.xml)
Response: raw XML (Content-Type: application/xml)

### GET /api/element?doc=<path>&file=<xmlPath>&xpath=<xpath>
XPath로 특정 요소를 조회. XML과 JSON 모두 반환.
- xpath: 슬래시 구분, 인덱스는 [N] (예: /hp:sec/hp:p[0]/hp:run[0]/hp:t)
Response: { "file": "...", "xpath": "...", "xml": "<요소XML>", "json": {...} }

### PUT /api/element?doc=<path>&file=<xmlPath>&xpath=<xpath>
XPath로 특정 요소를 수정. Body는 JSON, 3가지 방식 중 택 1:
- {"text": "새 텍스트"} — 요소 내 텍스트만 변경
- {"xml": "<hp:t>새 내용</hp:t>"} — XML 문자열로 요소 전체 교체
- {"json": {...}} — JSON 객체로 요소 교체
Response: { "success": true }

### POST /api/save?doc=<path>
수정사항을 HWPX 파일에 저장. 저장 후 뷰어가 자동으로 새로고침됩니다.
Response: { "success": true, "path": "..." }

### POST /api/reload?doc=<path>
뷰어를 새로고침합니다. 저장 없이 현재 메모리 상태를 반영.
Response: { "success": true, "path": "..." }

## XPath 형식
HWPX 문서는 XML 기반(OWPML KS X 6101:2024).

절대 경로 예시: /hs:sec/hp:p[0]/hp:run[0]/hp:t
- 루트 태그는 문서마다 다를 수 있음 (hs:sec, hp:sec 등). /api/xml로 확인 가능.
- 태그명에 네임스페이스 접두사 포함 (hp:, hs:, hh:, hc: 등)
- [N]은 같은 태그의 N번째 (0-based). 생략 시 [0]과 동일.
- VS Code에서 Select 모드로 요소 클릭 시 경로가 클립보드에 복사됨

### Descendant 검색 (//)
경로 중간에 //를 사용하면 하위 전체에서 검색합니다.
예: /hs:sec//hp:t[0] → 문서 내 첫 번째 hp:t 요소

### 상대 경로
/로 시작하지 않으면 상대 경로로, 루트에서 descendant 검색합니다.
예: hp:p[2]/hp:run[0]/hp:t → 문서 내 3번째 hp:p에서 시작

### Wrapper 자동 스킵
테이블 셀 내부의 hp:subList 같은 중간 래퍼 요소는 자동으로 스킵됩니다.
/hs:sec/hp:tbl/hp:tr[0]/hp:tc[0]/hp:p[0] 으로 요청해도
실제 구조가 hp:tc/hp:subList/hp:p인 경우 자동으로 찾아줍니다.

## 현재 열린 문서
${docs.length > 0 ? docs.map(d => `- ${d}`).join('\n') : '(없음)'}

## 사용 예시
\`\`\`bash
# 문서 목록
curl -H "Authorization: Bearer <token>" ${base}/api/documents

# 파일 목록
curl -H "Authorization: Bearer <token>" "${base}/api/files"

# XML 원본 조회
curl -H "Authorization: Bearer <token>" "${base}/api/xml?file=Contents/section0.xml"

# 요소 조회
curl -H "Authorization: Bearer <token>" "${base}/api/element?file=Contents/section0.xml&xpath=/hs:sec/hp:p[0]/hp:run[0]/hp:t"

# 텍스트 수정
curl -X PUT -H "Authorization: Bearer <token>" -H "Content-Type: application/json" \\
  "${base}/api/element?file=Contents/section0.xml&xpath=/hs:sec/hp:p[0]/hp:run[0]/hp:t" \\
  -d '{"text": "수정된 텍스트"}'

# 저장 (뷰어 자동 새로고침)
curl -X POST -H "Authorization: Bearer <token>" "${base}/api/save"

# 뷰어 새로고침 (저장 없이)
curl -X POST -H "Authorization: Bearer <token>" "${base}/api/reload"
\`\`\`
`;

        res.writeHead(200, {
            'Content-Type': 'text/markdown; charset=utf-8',
            'Content-Length': Buffer.byteLength(helpText, 'utf-8'),
            'Connection': 'close',
        });
        res.end(helpText);
    }

    /** GET /api/documents — 열려있는 HWPX 문서 목록 */
    private async handleListDocuments(res: http.ServerResponse) {
        const docs = Array.from(this.documents.entries()).map(([key, doc]) => ({
            path: key,
            dirty: doc.dirty,
        }));
        this.sendJson(res, 200, { documents: docs, port: this.port });
    }

    /** GET /api/files?doc=<path> — 문서 내 XML 파일 목록 */
    private async handleListFiles(url: URL, res: http.ServerResponse) {
        const doc = this.getDoc(url, res);
        if (!doc) return;

        const files: string[] = [];
        doc.zip.forEach((relativePath) => {
            if (relativePath.endsWith('.xml') || relativePath === 'mimetype') {
                files.push(relativePath);
            }
        });
        this.sendJson(res, 200, { files: files.sort() });
    }

    /** GET /api/xml?doc=<path>&file=<xmlPath> — XML 파일 전체 내용 */
    private async handleGetXml(url: URL, res: http.ServerResponse) {
        const doc = this.getDoc(url, res);
        if (!doc) return;

        const filePath = url.searchParams.get('file');
        if (!filePath) {
            this.sendJson(res, 400, { error: 'Missing "file" parameter' });
            return;
        }

        const file = doc.zip.file(filePath);
        if (!file) {
            this.sendJson(res, 404, { error: `File not found in HWPX: ${filePath}` });
            return;
        }

        const xml = await file.async('string');
        res.writeHead(200, {
            'Content-Type': 'application/xml; charset=utf-8',
            'Content-Length': Buffer.byteLength(xml, 'utf-8'),
            'Connection': 'close',
        });
        res.end(xml);
    }

    /** GET /api/element?doc=<path>&file=<xmlPath>&xpath=<xpath> — 특정 요소 조회 */
    private async handleGetElement(url: URL, res: http.ServerResponse) {
        const doc = this.getDoc(url, res);
        if (!doc) return;

        const filePath = url.searchParams.get('file');
        const xpath = url.searchParams.get('xpath');
        if (!filePath || !xpath) {
            this.sendJson(res, 400, { error: 'Missing "file" or "xpath" parameter' });
            return;
        }

        const file = doc.zip.file(filePath);
        if (!file) {
            this.sendJson(res, 404, { error: `File not found in HWPX: ${filePath}` });
            return;
        }

        const xml = await file.async('string');
        const parsed = this.parser.parse(xml);
        const element = this.navigateXpath(parsed, xpath);

        if (element === undefined) {
            this.sendJson(res, 404, { error: `Element not found at xpath: ${xpath}` });
            return;
        }

        // 요소를 다시 XML로 변환
        const elementXml = this.builder.build(Array.isArray(element) ? element : [element]);

        this.sendJson(res, 200, {
            file: filePath,
            xpath: xpath,
            xml: elementXml,
            json: element,
        });
    }

    /** PUT /api/element?doc=<path>&file=<xmlPath>&xpath=<xpath> — 요소 수정 */
    private async handlePutElement(url: URL, req: http.IncomingMessage, res: http.ServerResponse) {
        const doc = this.getDoc(url, res);
        if (!doc) return;

        const filePath = url.searchParams.get('file');
        const xpath = url.searchParams.get('xpath');
        if (!filePath || !xpath) {
            this.sendJson(res, 400, { error: 'Missing "file" or "xpath" parameter' });
            return;
        }

        const body = await this.readBody(req);
        let newContent: any;
        try {
            const bodyObj = JSON.parse(body);
            if (bodyObj.xml) {
                // XML 문자열로 수정
                newContent = { xml: bodyObj.xml };
            } else if (bodyObj.json) {
                // JSON 객체로 수정
                newContent = { json: bodyObj.json };
            } else if (bodyObj.text !== undefined) {
                // 텍스트 내용만 수정
                newContent = { text: bodyObj.text };
            } else {
                this.sendJson(res, 400, { error: 'Body must contain "xml", "json", or "text" field' });
                return;
            }
        } catch {
            this.sendJson(res, 400, { error: 'Invalid JSON body' });
            return;
        }

        const file = doc.zip.file(filePath);
        if (!file) {
            this.sendJson(res, 404, { error: `File not found in HWPX: ${filePath}` });
            return;
        }

        const xml = await file.async('string');
        const parsed = this.parser.parse(xml);

        // xpath로 부모와 인덱스 찾기
        const navResult = this.navigateXpathForUpdate(parsed, xpath);
        if (!navResult) {
            this.sendJson(res, 404, { error: `Element not found at xpath: ${xpath}` });
            return;
        }

        // 요소 교체
        if (newContent.xml) {
            const newParsed = this.parser.parse(newContent.xml);
            const newItem = newParsed[0] || newParsed;
            // preserveOrder: newItem은 {tagName: [...], ":@": {...}} 형태
            // parent[key]의 값(content)만 교체하고, 속성도 갱신
            const newTagName = Object.keys(newItem).find(k => k !== ':@');
            if (newTagName && newTagName === navResult.key) {
                navResult.parent[navResult.key] = newItem[newTagName];
                if (newItem[':@']) {
                    navResult.parent[':@'] = { ...(navResult.parent[':@'] || {}), ...newItem[':@'] };
                }
            } else {
                // 태그명이 다른 경우: 기존 키 제거 후 새 키 삽입
                navResult.parent[navResult.key] = newItem[newTagName || navResult.key] || newItem;
            }
        } else if (newContent.json) {
            navResult.parent[navResult.key] = newContent.json;
        } else if (newContent.text !== undefined) {
            // 텍스트만 수정: #text 노드 교체
            const target = navResult.index !== undefined
                ? navResult.parent[navResult.key][navResult.index]
                : navResult.parent[navResult.key];
            this.setTextContent(target, newContent.text);

            // 텍스트 변경 시 hp:linesegarray 제거 (레이아웃 메타데이터 무효화)
            this.removeLinesegarray(parsed, xpath);
        }

        // XML 빌드 (원본 whitespace 보존 — setTextContent가 hp:t만 정확히 수정)
        const newXml = this.builder.build(parsed);
        doc.zip.file(filePath, newXml);
        doc.dirty = true;

        this.sendJson(res, 200, { success: true, file: filePath, xpath: xpath });
    }

    /** POST /api/save?doc=<path> — 문서 저장 */
    private async handleSave(url: URL, res: http.ServerResponse) {
        const doc = this.getDoc(url, res);
        if (!doc) return;

        const content = await doc.zip.generateAsync({ type: 'uint8array' });
        await vscode.workspace.fs.writeFile(doc.uri, content);
        doc.dirty = false;

        // 저장 후 뷰어 자동 리로드
        this.triggerReload(doc.uri.fsPath);

        this.sendJson(res, 200, { success: true, path: doc.uri.fsPath });
    }

    /** POST /api/reload?doc=<path> — 뷰어 새로고침 */
    private async handleReload(url: URL, res: http.ServerResponse) {
        const doc = this.getDoc(url, res);
        if (!doc) return;

        this.triggerReload(doc.uri.fsPath);

        this.sendJson(res, 200, { success: true, path: doc.uri.fsPath });
    }

    // --- Helper methods ---

    private getDoc(url: URL, res: http.ServerResponse): OpenDocument | null {
        const docPath = url.searchParams.get('doc');
        if (!docPath) {
            // 문서가 하나만 열려있으면 자동 선택
            if (this.documents.size === 1) {
                return this.documents.values().next().value!;
            }
            this.sendJson(res, 400, { error: 'Missing "doc" parameter. Specify document path.', available: Array.from(this.documents.keys()) });
            return null;
        }
        const doc = this.documents.get(docPath);
        if (!doc) {
            this.sendJson(res, 404, { error: `Document not found: ${docPath}`, available: Array.from(this.documents.keys()) });
            return null;
        }
        return doc;
    }

    /**
     * XPath 탐색 (preserveOrder 모드 기준)
     * 지원 형식:
     *   /hs:sec/hp:p[0]/hp:run[0]/hp:t          — 절대 경로
     *   hp:p[0]/hp:run[0]/hp:t                    — 상대 경로 (루트에서 descendant 탐색)
     *   /hs:sec//hp:t                             — descendant 축 (//)
     *   /hs:sec/hp:tbl/hp:tr[0]/hp:tc[0]/hp:p[0] — 중간 wrapper 자동 스킵 (hp:subList 등)
     */
    private navigateXpath(obj: any, xpath: string): any {
        const segments = this.parseXpathSegments(xpath);
        if (segments.length === 0) return undefined;

        // 상대 경로: 첫 세그먼트가 descendant면 루트 전체에서 검색
        if (segments[0].type === 'descendant') {
            return this.findDescendant(obj, segments[0].part!, segments.slice(1));
        }

        // 절대 경로
        let current: any = obj;
        for (let i = 0; i < segments.length; i++) {
            const seg = segments[i];
            if (seg.type === 'descendant') {
                // // 이후의 나머지 경로
                const rest = segments.slice(i + 1);
                return this.findDescendant(current, seg.part!, rest);
            }
            current = this.stepInto(current, seg.part!);
            if (current === undefined) {
                // 한 단계 깊이 자동 스킵 시도 (wrapper 요소 관용)
                // 이전 위치로 돌아가서 모든 자식 중에서 찾기
                const prev = i === 0 ? obj : this.navigateToIndex(obj, segments, i);
                if (prev !== undefined) {
                    const skipped = this.trySkipWrapper(prev, seg.part!);
                    if (skipped !== undefined) {
                        current = skipped;
                        continue;
                    }
                }
                return undefined;
            }
        }
        return current;
    }

    /**
     * 업데이트를 위해 부모와 키/인덱스를 반환
     */
    private navigateXpathForUpdate(obj: any, xpath: string): { parent: any; key: string; index?: number } | null {
        const segments = this.parseXpathSegments(xpath);
        if (segments.length === 0) return null;

        // descendant(//), 상대 경로는 먼저 전체 경로로 해석하여 대상 찾기
        const target = this.navigateXpath(obj, xpath);
        if (target === undefined) return null;

        // 대상을 찾은 후, 부모를 역추적하여 반환
        // 마지막 세그먼트의 part로 부모를 찾는다
        const lastSeg = segments[segments.length - 1];
        const lastPart = lastSeg.type === 'descendant' ? lastSeg.part! : lastSeg.part!;

        // 부모 경로로 탐색
        const parentSegments = segments.slice(0, -1);
        let parent: any;
        if (parentSegments.length === 0) {
            parent = obj;
        } else {
            // 부모까지의 경로 재구성
            const parentXpath = this.rebuildXpath(parentSegments);
            parent = this.navigateXpath(obj, parentXpath);
        }

        if (parent === undefined) return null;

        // preserveOrder 배열에서 부모 찾기
        if (Array.isArray(parent)) {
            let count = 0;
            for (const item of parent) {
                if (item[lastPart.name] !== undefined) {
                    if (count === lastPart.index) {
                        return { parent: item, key: lastPart.name };
                    }
                    count++;
                }
            }
            // wrapper 스킵하여 찾기
            for (const item of parent) {
                for (const key of Object.keys(item)) {
                    if (key.startsWith(':@') || key === '#text') continue;
                    const child = item[key];
                    if (Array.isArray(child)) {
                        let innerCount = 0;
                        for (const innerItem of child) {
                            if (innerItem[lastPart.name] !== undefined) {
                                if (innerCount === lastPart.index) {
                                    return { parent: innerItem, key: lastPart.name };
                                }
                                innerCount++;
                            }
                        }
                    }
                }
            }
        } else if (typeof parent === 'object') {
            if (parent[lastPart.name] !== undefined) {
                if (Array.isArray(parent[lastPart.name]) && lastPart.index > 0) {
                    return { parent, key: lastPart.name, index: lastPart.index };
                }
                return { parent, key: lastPart.name };
            }
            // wrapper 스킵
            for (const key of Object.keys(parent)) {
                if (key.startsWith(':@') || key === '#text') continue;
                const child = parent[key];
                if (typeof child === 'object' && child !== null) {
                    const arr = Array.isArray(child) ? child : [child];
                    for (const item of arr) {
                        if (item[lastPart.name] !== undefined) {
                            return { parent: item, key: lastPart.name };
                        }
                    }
                }
            }
        }

        return null;
    }

    private rebuildXpath(segments: XpathSegment[]): string {
        let result = '';
        for (const seg of segments) {
            if (seg.type === 'descendant') {
                result += `//${seg.part!.name}` + (seg.part!.index > 0 ? `[${seg.part!.index}]` : '');
            } else {
                result += `/${seg.part!.name}` + (seg.part!.index > 0 ? `[${seg.part!.index}]` : '');
            }
        }
        return result;
    }

    /** preserveOrder 배열/객체에서 한 단계 이동 */
    private stepInto(current: any, part: { name: string; index: number }): any {
        if (Array.isArray(current)) {
            let count = 0;
            for (const item of current) {
                if (item[part.name] !== undefined) {
                    if (count === part.index) {
                        return item[part.name];
                    }
                    count++;
                }
            }
            return undefined;
        } else if (typeof current === 'object' && current !== null) {
            const val = current[part.name];
            if (val === undefined) return undefined;
            if (Array.isArray(val) && part.index > 0) {
                return val[part.index];
            }
            return val;
        }
        return undefined;
    }

    /** 인덱스 i까지의 경로를 다시 탐색 */
    private navigateToIndex(obj: any, segments: XpathSegment[], targetIdx: number): any {
        let current: any = obj;
        for (let i = 0; i < targetIdx; i++) {
            const seg = segments[i];
            if (seg.type === 'descendant' || !seg.part) return undefined;
            current = this.stepInto(current, seg.part);
            if (current === undefined) return undefined;
        }
        return current;
    }

    /** wrapper 요소(hp:subList 등)를 1단계 스킵하여 대상 찾기 */
    private trySkipWrapper(parent: any, part: { name: string; index: number }): any {
        const items = Array.isArray(parent) ? parent : [parent];
        for (const item of items) {
            if (typeof item !== 'object' || item === null) continue;
            for (const key of Object.keys(item)) {
                if (key.startsWith(':@') || key === '#text') continue;
                const child = item[key];
                if (child === undefined) continue;
                // child 안에서 대상 탐색
                const result = this.stepInto(child, part);
                if (result !== undefined) return result;
            }
        }
        return undefined;
    }

    /** 재귀적으로 descendant 검색 */
    private findDescendant(obj: any, part: { name: string; index: number }, rest: XpathSegment[]): any {
        const matches: any[] = [];
        this.collectByName(obj, part.name, matches);
        if (part.index >= matches.length) return undefined;
        const found = matches[part.index];
        if (rest.length === 0) return found;
        // 나머지 경로 계속 탐색
        let current = found;
        for (let ri = 0; ri < rest.length; ri++) {
            const seg = rest[ri];
            if (seg.type === 'descendant') {
                return this.findDescendant(current, seg.part!, rest.slice(ri + 1));
            }
            const prev = current;
            current = this.stepInto(current, seg.part!);
            if (current === undefined) {
                // wrapper 스킵 시도: prev의 자식들을 한 단계 더 깊이 탐색
                const skipped = this.trySkipWrapper(prev, seg.part!);
                if (skipped !== undefined) {
                    current = skipped;
                } else {
                    return undefined;
                }
            }
        }
        return current;
    }

    /** 이름으로 모든 descendant 수집 (preserveOrder: 값을 배열째 push) */
    private collectByName(obj: any, name: string, results: any[]) {
        if (Array.isArray(obj)) {
            for (const item of obj) {
                if (typeof item === 'object' && item !== null) {
                    if (item[name] !== undefined) {
                        // preserveOrder에서 값은 항상 배열 — 배열째 하나의 매치로 push
                        results.push(item[name]);
                    }
                    for (const key of Object.keys(item)) {
                        if (key === name || key.startsWith(':@') || key === '#text' || key === '?xml') continue;
                        this.collectByName(item[key], name, results);
                    }
                }
            }
        } else if (typeof obj === 'object' && obj !== null) {
            if (obj[name] !== undefined) {
                results.push(obj[name]);
            }
            for (const key of Object.keys(obj)) {
                if (key === name || key.startsWith(':@') || key === '#text' || key === '?xml') continue;
                this.collectByName(obj[key], name, results);
            }
        }
    }

    private parseXpathSegments(xpath: string): XpathSegment[] {
        // "/hs:sec/hp:p[0]//hp:t" → segments
        // "hp:p[0]/hp:run[0]" (상대 경로) → 첫 세그먼트를 descendant로
        const trimmed = xpath.trim();
        const isAbsolute = trimmed.startsWith('/');
        const segments: XpathSegment[] = [];

        // // 로 분할
        const doubleSlashParts = trimmed.split('//');
        for (let di = 0; di < doubleSlashParts.length; di++) {
            const chunk = doubleSlashParts[di];
            if (!chunk && di === 0) continue; // 맨 앞 / 제거
            const singleParts = chunk.split('/').filter(Boolean);

            for (let si = 0; si < singleParts.length; si++) {
                const part = this.parseSegment(singleParts[si]);
                if (di > 0 && si === 0) {
                    // // 바로 뒤의 첫 세그먼트는 descendant
                    segments.push({ type: 'descendant', part });
                } else if (!isAbsolute && segments.length === 0) {
                    // 상대 경로의 첫 세그먼트 → descendant로 처리
                    segments.push({ type: 'descendant', part });
                } else {
                    segments.push({ type: 'child', part });
                }
            }
        }
        return segments;
    }

    private parseSegment(segment: string): { name: string; index: number } {
        const match = segment.match(/^(.+?)\[(\d+)\]$/);
        if (match) {
            return { name: match[1], index: parseInt(match[2], 10) };
        }
        return { name: segment, index: 0 };
    }

    /**
     * 텍스트 변경 시, xpath에서 가장 가까운 hp:p 조상을 찾아
     * hp:linesegarray를 제거한다 (레이아웃 메타데이터 무효화).
     * 뷰어/한글이 문서를 열 때 자동으로 재계산한다.
     */
    private removeLinesegarray(parsed: any, xpath: string) {
        const segments = this.parseXpathSegments(xpath);

        // xpath에서 마지막 hp:p 세그먼트까지의 경로를 찾아 해당 hp:p 배열을 얻는다
        let lastPIdx = -1;
        for (let i = 0; i < segments.length; i++) {
            if (segments[i].part && segments[i].part.name === 'hp:p') {
                lastPIdx = i;
            }
        }
        if (lastPIdx < 0) return;

        // hp:p까지 탐색
        const pXpath = this.rebuildXpath(segments.slice(0, lastPIdx + 1));
        const pElement = this.navigateXpath(parsed, pXpath);
        if (!pElement || !Array.isArray(pElement)) return;

        // hp:linesegarray 항목 제거
        for (let i = pElement.length - 1; i >= 0; i--) {
            if (pElement[i] && pElement[i]['hp:linesegarray'] !== undefined) {
                pElement.splice(i, 1);
            }
        }
    }

    /**
     * 요소 내 텍스트를 수정한다.
     * hp:t 요소의 #text만 대상으로 하여 whitespace #text 노드를 건드리지 않는다.
     * hp:run, hp:br 등 기존 서식 구조는 보존된다.
     */
    private setTextContent(obj: any, text: string): boolean {
        // 전략 1: hp:t 자손을 찾아 그 안의 #text를 교체
        if (this.replaceHpTText(obj, text)) return true;

        // 전략 2: hp:t가 없는 경우 (대상이 hp:t 자체이거나 단순 구조)
        // whitespace-only #text를 건너뛰고 실제 콘텐츠 #text를 교체
        return this.replaceFirstContentText(obj, text);
    }

    /** hp:t 요소를 재귀 탐색하여 첫 번째 hp:t의 #text를 교체 */
    private replaceHpTText(obj: any, text: string): boolean {
        if (Array.isArray(obj)) {
            for (const item of obj) {
                if (typeof item !== 'object' || item === null) continue;
                if (item['hp:t'] !== undefined) {
                    // hp:t 내부의 #text 교체
                    const hpT = item['hp:t'];
                    if (Array.isArray(hpT)) {
                        for (const tItem of hpT) {
                            if (tItem && '#text' in tItem) {
                                tItem['#text'] = text;
                                return true;
                            }
                        }
                    }
                    continue;
                }
                // 자식 요소로 재귀 (whitespace #text, 속성 스킵)
                for (const key of Object.keys(item)) {
                    if (key === '#text' || key.startsWith(':@')) continue;
                    if (this.replaceHpTText(item[key], text)) return true;
                }
            }
        } else if (typeof obj === 'object' && obj !== null) {
            if (obj['hp:t'] !== undefined) {
                const hpT = obj['hp:t'];
                if (Array.isArray(hpT)) {
                    for (const tItem of hpT) {
                        if (tItem && '#text' in tItem) {
                            tItem['#text'] = text;
                            return true;
                        }
                    }
                }
                return false;
            }
            for (const key of Object.keys(obj)) {
                if (key === '#text' || key.startsWith(':@')) continue;
                if (this.replaceHpTText(obj[key], text)) return true;
            }
        }
        return false;
    }

    /** whitespace-only가 아닌 첫 번째 #text를 교체 (hp:t가 없는 경우 fallback) */
    private replaceFirstContentText(obj: any, text: string): boolean {
        if (Array.isArray(obj)) {
            for (const item of obj) {
                if (typeof item !== 'object' || item === null) continue;
                if ('#text' in item && typeof item['#text'] === 'string' && item['#text'].trim() !== '') {
                    item['#text'] = text;
                    return true;
                }
                for (const key of Object.keys(item)) {
                    if (key === '#text' || key.startsWith(':@')) continue;
                    if (this.replaceFirstContentText(item[key], text)) return true;
                }
            }
        } else if (typeof obj === 'object' && obj !== null) {
            if ('#text' in obj && typeof obj['#text'] === 'string' && obj['#text'].trim() !== '') {
                obj['#text'] = text;
                return true;
            }
            for (const key of Object.keys(obj)) {
                if (key === '#text' || key.startsWith(':@')) continue;
                if (this.replaceFirstContentText(obj[key], text)) return true;
            }
        }
        return false;
    }

    private sendJson(res: http.ServerResponse, status: number, data: any) {
        const body = JSON.stringify(data, null, 2);
        res.writeHead(status, {
            'Content-Type': 'application/json; charset=utf-8',
            'Content-Length': Buffer.byteLength(body, 'utf-8'),
            'Connection': 'close',
        });
        res.end(body);
    }

    private readBody(req: http.IncomingMessage): Promise<string> {
        return new Promise((resolve, reject) => {
            const chunks: Buffer[] = [];
            req.on('data', (chunk: Buffer) => chunks.push(chunk));
            req.on('end', () => resolve(Buffer.concat(chunks).toString('utf-8')));
            req.on('error', reject);
        });
    }
}
