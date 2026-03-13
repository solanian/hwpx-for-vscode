/**
 * HwpxApiServer 단위 테스트
 * 실행: npm test
 */

// vscode mock — 반드시 HwpxApiServer import 전에 설정
let lastWrittenUri: any = null;
let lastWrittenContent: any = null;

const mockVscode = {
    Uri: {
        file: (path: string) => ({ fsPath: path, scheme: 'file', path }),
    },
    workspace: {
        fs: {
            writeFile: async (uri: any, content: Uint8Array) => {
                lastWrittenUri = uri;
                lastWrittenContent = content;
            },
        },
    },
    window: {
        showInformationMessage: () => {},
        showWarningMessage: () => {},
    },
    commands: {
        registerCommand: () => ({ dispose: () => {} }),
    },
    env: {
        clipboard: { writeText: async () => {} },
    },
};

// Module mock 주입
import Module from 'module';
const originalRequire = (Module as any).prototype.require;
(Module as any).prototype.require = function (id: string) {
    if (id === 'vscode') return mockVscode;
    return originalRequire.apply(this, arguments);
};

import * as http from 'http';
import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';
import { HwpxApiServer } from '../src/hwpxApiServer';

// --- Test helpers ---

let passed = 0;
let failed = 0;
const errors: string[] = [];

function assert(condition: boolean, msg: string) {
    if (!condition) {
        failed++;
        errors.push(`FAIL: ${msg}`);
        console.error(`  ✗ ${msg}`);
    } else {
        passed++;
        console.log(`  ✓ ${msg}`);
    }
}

function assertEqual(actual: any, expected: any, msg: string) {
    if (actual !== expected) {
        failed++;
        const detail = `${msg} — expected: ${JSON.stringify(expected)}, got: ${JSON.stringify(actual)}`;
        errors.push(`FAIL: ${detail}`);
        console.error(`  ✗ ${detail}`);
    } else {
        passed++;
        console.log(`  ✓ ${msg}`);
    }
}

function assertIncludes(str: string, substr: string, msg: string) {
    if (!str || !str.includes(substr)) {
        failed++;
        const detail = `${msg} — "${substr}" not found in response`;
        errors.push(`FAIL: ${detail}`);
        console.error(`  ✗ ${detail}`);
    } else {
        passed++;
        console.log(`  ✓ ${msg}`);
    }
}

function assertNotEqual(actual: any, expected: any, msg: string) {
    if (actual === expected) {
        failed++;
        const detail = `${msg} — should not be: ${JSON.stringify(expected)}`;
        errors.push(`FAIL: ${detail}`);
        console.error(`  ✗ ${detail}`);
    } else {
        passed++;
        console.log(`  ✓ ${msg}`);
    }
}

interface HttpResult {
    status: number;
    headers: http.IncomingHttpHeaders;
    body: string;
    json: any;
}

function httpRequest(port: number, method: string, urlPath: string, body?: string, headers?: Record<string, string>): Promise<HttpResult> {
    return new Promise((resolve, reject) => {
        const reqHeaders: Record<string, string> = { ...(headers || {}) };
        if (body && !reqHeaders['Content-Type']) {
            reqHeaders['Content-Type'] = 'application/json';
        }
        const options: http.RequestOptions = {
            hostname: '127.0.0.1',
            port,
            path: urlPath,
            method,
            headers: reqHeaders,
        };
        const req = http.request(options, (res) => {
            const chunks: Buffer[] = [];
            res.on('data', (chunk: Buffer) => chunks.push(chunk));
            res.on('end', () => {
                const bodyStr = Buffer.concat(chunks).toString('utf-8');
                let json: any = null;
                try { json = JSON.parse(bodyStr); } catch {}
                resolve({ status: res.statusCode || 0, headers: res.headers, body: bodyStr, json });
            });
        });
        req.on('error', reject);
        if (body) req.write(body);
        req.end();
    });
}

/** 토큰 포함된 요청 helper */
function authRequest(port: number, token: string, method: string, urlPath: string, body?: string, headers?: Record<string, string>): Promise<HttpResult> {
    return httpRequest(port, method, urlPath, body, {
        ...headers,
        'Authorization': `Bearer ${token}`,
    });
}

// --- Main test ---

function createMockZip(): JSZip {
    const zip = new JSZip();
    zip.file('mimetype', 'application/hwp+zip');
    zip.file('Contents/content.hpf', `<?xml version="1.0" encoding="UTF-8"?>
<opf:package xmlns:opf="http://www.idpf.org/2007/opf">
  <opf:metadata>
    <opf:title>테스트 문서</opf:title>
  </opf:metadata>
</opf:package>`);
    zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hp:sec xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p>
    <hp:run>
      <hp:t>안녕하세요 테스트입니다</hp:t>
    </hp:run>
  </hp:p>
  <hp:p>
    <hp:run>
      <hp:t>두번째 문단</hp:t>
    </hp:run>
    <hp:run>
      <hp:t>두번째 런</hp:t>
    </hp:run>
  </hp:p>
</hp:sec>`);
    zip.file('Contents/header.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:mappingTable>
    <hh:charPr id="0" height="1000"/>
  </hh:mappingTable>
</hh:head>`);
    return zip;
}

async function runTests() {
    console.log('\n=== HwpxApiServer Unit Tests ===\n');

    const server = new HwpxApiServer();
    const port = await server.start();
    const token = server.getToken();
    console.log(`Server started on port ${port}`);
    console.log(`Token: ${token.substring(0, 8)}...\n`);

    const zip = createMockZip();
    const fakeUri = mockVscode.Uri.file('/test/doc.hwpx') as any;
    server.registerDocument(fakeUri, zip);

    const FILE = 'Contents/section0.xml';
    const ROOT = 'hp:sec';

    // ============================================================
    // AUTH: 토큰 인증
    // ============================================================
    console.log('--- Auth: token required ---');
    {
        // 토큰 없이 요청 → 401
        const res = await httpRequest(port, 'GET', '/api/documents');
        assertEqual(res.status, 401, 'auth: no token returns 401');
        assertIncludes(res.json.error, 'Unauthorized', 'auth: error message');
    }
    {
        // 잘못된 토큰 → 401
        const res = await httpRequest(port, 'GET', '/api/documents', undefined, {
            'Authorization': 'Bearer wrong-token',
        });
        assertEqual(res.status, 401, 'auth: wrong token returns 401');
    }
    {
        // Authorization 헤더로 인증
        const res = await authRequest(port, token, 'GET', '/api/documents');
        assertEqual(res.status, 200, 'auth: Bearer token accepted');
    }
    {
        // ?token= 쿼리 파라미터로 인증
        const res = await httpRequest(port, 'GET', `/api/documents?token=${token}`);
        assertEqual(res.status, 200, 'auth: query token accepted');
    }
    {
        // OPTIONS는 토큰 없이도 OK (CORS preflight)
        const res = await httpRequest(port, 'OPTIONS', '/api/documents');
        assertEqual(res.status, 200, 'auth: OPTIONS without token returns 200');
    }

    // ============================================================
    // CORS: Origin 제한
    // ============================================================
    console.log('\n--- CORS: origin restrictions ---');
    {
        // Origin 없음 (curl 등) → 허용
        const res = await authRequest(port, token, 'GET', '/api/documents');
        assertEqual(res.status, 200, 'cors: no origin allowed');
        assert(res.headers['access-control-allow-origin'] !== undefined, 'cors: allow-origin header present');
    }
    {
        // localhost origin → 허용
        const res = await httpRequest(port, 'GET', `/api/documents?token=${token}`, undefined, {
            'Origin': 'http://127.0.0.1:3000',
        });
        assertEqual(res.status, 200, 'cors: localhost origin allowed');
        assertEqual(res.headers['access-control-allow-origin'], 'http://127.0.0.1:3000', 'cors: reflects localhost origin');
    }
    {
        // vscode-webview origin → 허용
        const res = await httpRequest(port, 'GET', `/api/documents?token=${token}`, undefined, {
            'Origin': 'vscode-webview://abc123',
        });
        assertEqual(res.headers['access-control-allow-origin'], 'vscode-webview://abc123', 'cors: reflects vscode-webview origin');
    }
    {
        // 외부 origin → 거부 (null)
        const res = await httpRequest(port, 'GET', `/api/documents?token=${token}`, undefined, {
            'Origin': 'https://evil.com',
        });
        assertEqual(res.headers['access-control-allow-origin'], 'null', 'cors: external origin blocked');
    }
    {
        // Authorization 헤더가 허용 헤더에 포함되는지
        const res = await httpRequest(port, 'OPTIONS', '/api/documents', undefined, {
            'Origin': 'http://127.0.0.1:3000',
        });
        assertIncludes(res.headers['access-control-allow-headers'] || '', 'Authorization', 'cors: Authorization in allowed headers');
    }

    // ============================================================
    // GET /api/documents
    // ============================================================
    console.log('\n--- GET /api/documents ---');
    {
        const res = await authRequest(port, token, 'GET', '/api/documents');
        assertEqual(res.status, 200, 'documents: status 200');
        assert(Array.isArray(res.json.documents), 'documents: returns array');
        assertEqual(res.json.documents.length, 1, 'documents: 1 document');
        assertEqual(res.json.documents[0].path, '/test/doc.hwpx', 'documents: correct path');
        assertEqual(res.json.documents[0].dirty, false, 'documents: not dirty');
    }

    // ============================================================
    // GET /api/files
    // ============================================================
    console.log('\n--- GET /api/files ---');
    {
        const res = await authRequest(port, token, 'GET', '/api/files');
        assertEqual(res.status, 200, 'files: status 200');
        assert(Array.isArray(res.json.files), 'files: returns array');
        const hasSection = res.json.files.some((f: string) => f.includes('section0.xml'));
        assert(hasSection, 'files: contains section0.xml');
        console.log(`  (${res.json.files.length} files: ${res.json.files.join(', ')})`);
    }

    // ============================================================
    // GET /api/xml
    // ============================================================
    console.log('\n--- GET /api/xml ---');
    {
        const res = await authRequest(port, token, 'GET', `/api/xml?file=${encodeURIComponent(FILE)}`);
        assertEqual(res.status, 200, 'xml: status 200');
        assertIncludes(res.headers['content-type'] || '', 'application/xml', 'xml: correct content-type');
        assertIncludes(res.body, 'hp:sec', 'xml: contains root element');
        assertIncludes(res.body, '안녕하세요', 'xml: contains original text');
    }
    {
        const res = await authRequest(port, token, 'GET', '/api/xml?file=nonexistent.xml');
        assertEqual(res.status, 404, 'xml missing file: status 404');
    }
    {
        const res = await authRequest(port, token, 'GET', '/api/xml');
        assertEqual(res.status, 400, 'xml no param: status 400');
    }

    // ============================================================
    // GET /api/element
    // ============================================================
    console.log('\n--- GET /api/element ---');
    {
        // 루트 요소
        const res = await authRequest(port, token, 'GET',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=${encodeURIComponent(`/${ROOT}`)}`);
        assertEqual(res.status, 200, 'element root: status 200');
        assert(res.json.xml !== undefined, 'element root: has xml');
        assert(res.json.json !== undefined, 'element root: has json');
        assertIncludes(res.json.xml, 'hp:', 'element root: xml has namespace');
    }
    {
        // 중첩 경로: 첫 번째 문단
        const res = await authRequest(port, token, 'GET',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=${encodeURIComponent(`/${ROOT}/hp:p[0]`)}`);
        assertEqual(res.status, 200, 'element p[0]: status 200');
        assertIncludes(res.json.xml, '안녕하세요', 'element p[0]: has first paragraph text');
    }
    {
        // 중첩 경로: 두 번째 문단
        const res = await authRequest(port, token, 'GET',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=${encodeURIComponent(`/${ROOT}/hp:p[1]`)}`);
        assertEqual(res.status, 200, 'element p[1]: status 200');
        assertIncludes(res.json.xml, '두번째 문단', 'element p[1]: has second paragraph text');
    }
    {
        // 깊은 경로: hp:t
        const res = await authRequest(port, token, 'GET',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=${encodeURIComponent(`/${ROOT}/hp:p[0]/hp:run[0]/hp:t`)}`);
        assertEqual(res.status, 200, 'element hp:t: status 200');
        assertIncludes(res.json.xml, '안녕하세요 테스트입니다', 'element hp:t: exact text');
    }
    {
        // 두 번째 런의 hp:t
        const res = await authRequest(port, token, 'GET',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=${encodeURIComponent(`/${ROOT}/hp:p[1]/hp:run[1]/hp:t`)}`);
        assertEqual(res.status, 200, 'element run[1]/t: status 200');
        assertIncludes(res.json.xml, '두번째 런', 'element run[1]/t: correct text');
    }
    {
        // 존재하지 않는 경로
        const res = await authRequest(port, token, 'GET',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=${encodeURIComponent('/nonexistent/path')}`);
        assertEqual(res.status, 404, 'element not found: status 404');
    }
    {
        // 파라미터 누락: file 없음
        const res = await authRequest(port, token, 'GET',
            `/api/element?xpath=${encodeURIComponent('/hp:sec')}`);
        assertEqual(res.status, 400, 'element no file param: status 400');
    }
    {
        // 파라미터 누락: xpath 없음
        const res = await authRequest(port, token, 'GET',
            `/api/element?file=${encodeURIComponent(FILE)}`);
        assertEqual(res.status, 400, 'element no xpath param: status 400');
    }

    // ============================================================
    // Descendant 검색 (//)
    // ============================================================
    console.log('\n--- Descendant search (//) ---');
    {
        // //hp:t[0] → 문서 내 첫 번째 hp:t
        const res = await authRequest(port, token, 'GET',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=${encodeURIComponent(`/${ROOT}//hp:t[0]`)}`);
        assertEqual(res.status, 200, 'descendant //hp:t[0]: status 200');
        assertIncludes(res.json.xml, '안녕하세요', 'descendant //hp:t[0]: first text');
    }
    {
        // //hp:t[2] → 세 번째 hp:t ("두번째 런")
        const res = await authRequest(port, token, 'GET',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=${encodeURIComponent(`/${ROOT}//hp:t[2]`)}`);
        assertEqual(res.status, 200, 'descendant //hp:t[2]: status 200');
        assertIncludes(res.json.xml, '두번째 런', 'descendant //hp:t[2]: third text');
    }
    {
        // //hp:run[0] → 첫 번째 런
        const res = await authRequest(port, token, 'GET',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=${encodeURIComponent(`/${ROOT}//hp:run[0]`)}`);
        assertEqual(res.status, 200, 'descendant //hp:run[0]: status 200');
    }

    // ============================================================
    // 상대 경로 (Relative path)
    // ============================================================
    console.log('\n--- Relative path ---');
    {
        // hp:p[0] → 루트에서 descendant로 첫 번째 hp:p 검색
        const res = await authRequest(port, token, 'GET',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=${encodeURIComponent('hp:p[0]/hp:run[0]/hp:t')}`);
        assertEqual(res.status, 200, 'relative hp:p[0]/hp:run[0]/hp:t: status 200');
        assertIncludes(res.json.xml, '안녕하세요', 'relative: correct text');
    }
    {
        // hp:t[1] → 두 번째 hp:t
        const res = await authRequest(port, token, 'GET',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=${encodeURIComponent('hp:t[1]')}`);
        assertEqual(res.status, 200, 'relative hp:t[1]: status 200');
        assertIncludes(res.json.xml, '두번째 문단', 'relative hp:t[1]: second text');
    }

    // ============================================================
    // Wrapper 자동 스킵 (hp:subList 등)
    // ============================================================
    console.log('\n--- Wrapper skip (subList) ---');
    {
        // subList가 포함된 mock 문서로 테스트
        const zipSub = new JSZip();
        zipSub.file('mimetype', 'application/hwp+zip');
        zipSub.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:tbl>
    <hp:tr>
      <hp:tc>
        <hp:subList>
          <hp:p><hp:run><hp:t>셀 내용</hp:t></hp:run></hp:p>
        </hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
</hs:sec>`);
        const subUri = mockVscode.Uri.file('/test/sub.hwpx') as any;
        server.registerDocument(subUri, zipSub);

        const subFile = 'Contents/section0.xml';
        const subDoc = encodeURIComponent('/test/sub.hwpx');

        // subList 포함한 정확한 경로
        const res1 = await authRequest(port, token, 'GET',
            `/api/element?doc=${subDoc}&file=${encodeURIComponent(subFile)}&xpath=${encodeURIComponent('/hs:sec/hp:tbl/hp:tr/hp:tc/hp:subList/hp:p/hp:run/hp:t')}`);
        assertEqual(res1.status, 200, 'wrapper exact: found with full path');
        assertIncludes(res1.json.xml, '셀 내용', 'wrapper exact: correct text');

        // subList 생략한 경로 → 자동 스킵
        const res2 = await authRequest(port, token, 'GET',
            `/api/element?doc=${subDoc}&file=${encodeURIComponent(subFile)}&xpath=${encodeURIComponent('/hs:sec/hp:tbl/hp:tr/hp:tc/hp:p/hp:run/hp:t')}`);
        assertEqual(res2.status, 200, 'wrapper skip: found without subList');
        assertIncludes(res2.json.xml, '셀 내용', 'wrapper skip: correct text');

        // descendant로도 찾기
        const res3 = await authRequest(port, token, 'GET',
            `/api/element?doc=${subDoc}&file=${encodeURIComponent(subFile)}&xpath=${encodeURIComponent('/hs:sec//hp:t')}`);
        assertEqual(res3.status, 200, 'wrapper descendant: found via //');
        assertIncludes(res3.json.xml, '셀 내용', 'wrapper descendant: correct text');

        server.unregisterDocument(subUri);
    }

    // ============================================================
    // PUT /api/element — text 수정
    // ============================================================
    console.log('\n--- PUT /api/element (text) ---');
    {
        const xpath = `/${ROOT}/hp:p[0]/hp:run[0]/hp:t`;
        const putRes = await authRequest(port, token, 'PUT',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=${encodeURIComponent(xpath)}`,
            JSON.stringify({ text: '수정된 텍스트' }));
        assertEqual(putRes.status, 200, 'put text: status 200');
        assert(putRes.json.success === true, 'put text: success');

        // 수정 확인
        const getRes = await authRequest(port, token, 'GET',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=${encodeURIComponent(xpath)}`);
        assertIncludes(getRes.json.xml, '수정된 텍스트', 'put text: verified');

        // dirty 확인
        const docsRes = await authRequest(port, token, 'GET', '/api/documents');
        assertEqual(docsRes.json.documents[0].dirty, true, 'put text: dirty=true');
    }

    // ============================================================
    // PUT /api/element — xml 수정
    // ============================================================
    console.log('\n--- PUT /api/element (xml) ---');
    {
        const xpath = `/${ROOT}/hp:p[0]/hp:run[0]/hp:t`;
        const newXml = '<hp:t xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">XML로 수정됨</hp:t>';
        const putRes = await authRequest(port, token, 'PUT',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=${encodeURIComponent(xpath)}`,
            JSON.stringify({ xml: newXml }));
        assertEqual(putRes.status, 200, 'put xml: status 200');
        assert(putRes.json.success === true, 'put xml: success');

        // 수정 확인
        const getRes = await authRequest(port, token, 'GET',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=${encodeURIComponent(xpath)}`);
        assertIncludes(getRes.json.xml, 'XML로 수정됨', 'put xml: verified');
    }

    // ============================================================
    // PUT /api/element — json 수정
    // ============================================================
    console.log('\n--- PUT /api/element (json) ---');
    {
        const xpath = `/${ROOT}/hp:p[0]/hp:run[0]/hp:t`;
        const getRes = await authRequest(port, token, 'GET',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=${encodeURIComponent(xpath)}`);
        const currentJson = getRes.json.json;

        // JSON 내에서 #text 수정
        const modified = JSON.parse(JSON.stringify(currentJson));
        if (Array.isArray(modified)) {
            for (const item of modified) {
                if (item['#text'] !== undefined) {
                    item['#text'] = 'JSON으로 수정됨';
                }
            }
        }

        const putRes = await authRequest(port, token, 'PUT',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=${encodeURIComponent(xpath)}`,
            JSON.stringify({ json: modified }));
        assertEqual(putRes.status, 200, 'put json: status 200');

        const verifyRes = await authRequest(port, token, 'GET',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=${encodeURIComponent(xpath)}`);
        assertIncludes(verifyRes.json.xml, 'JSON으로 수정됨', 'put json: verified');
    }

    // ============================================================
    // PUT /api/element — error cases
    // ============================================================
    console.log('\n--- PUT /api/element (errors) ---');
    {
        // Invalid JSON body
        const res = await authRequest(port, token, 'PUT',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=/${ROOT}`,
            'not json');
        assertEqual(res.status, 400, 'put error: invalid JSON → 400');
        assertIncludes(res.json.error, 'Invalid JSON', 'put error: correct message');
    }
    {
        // Body without xml/json/text
        const res = await authRequest(port, token, 'PUT',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=/${ROOT}`,
            JSON.stringify({ foo: 'bar' }));
        assertEqual(res.status, 400, 'put error: no xml/json/text → 400');
        assertIncludes(res.json.error, 'must contain', 'put error: correct message');
    }
    {
        // Missing file param
        const res = await authRequest(port, token, 'PUT',
            `/api/element?xpath=/${ROOT}`,
            JSON.stringify({ text: 'test' }));
        assertEqual(res.status, 400, 'put error: no file → 400');
    }
    {
        // Missing xpath param
        const res = await authRequest(port, token, 'PUT',
            `/api/element?file=${encodeURIComponent(FILE)}`,
            JSON.stringify({ text: 'test' }));
        assertEqual(res.status, 400, 'put error: no xpath → 400');
    }
    {
        // Non-existent file
        const res = await authRequest(port, token, 'PUT',
            `/api/element?file=nope.xml&xpath=/${ROOT}`,
            JSON.stringify({ text: 'test' }));
        assertEqual(res.status, 404, 'put error: missing file → 404');
    }
    {
        // Non-existent xpath
        const res = await authRequest(port, token, 'PUT',
            `/api/element?file=${encodeURIComponent(FILE)}&xpath=/no/such/path`,
            JSON.stringify({ text: 'test' }));
        assertEqual(res.status, 404, 'put error: missing xpath → 404');
    }

    // ============================================================
    // POST /api/save
    // ============================================================
    console.log('\n--- POST /api/save ---');
    {
        lastWrittenUri = null;
        lastWrittenContent = null;

        const res = await authRequest(port, token, 'POST', '/api/save');
        assertEqual(res.status, 200, 'save: status 200');
        assert(res.json.success === true, 'save: success');
        assertIncludes(res.json.path, 'doc.hwpx', 'save: correct path');

        // mock writeFile 호출 확인
        assert(lastWrittenUri !== null, 'save: writeFile called');
        assert(lastWrittenContent !== null && lastWrittenContent.length > 0, 'save: content written');

        // dirty → false 확인
        const docsRes = await authRequest(port, token, 'GET', '/api/documents');
        assertEqual(docsRes.json.documents[0].dirty, false, 'save: dirty=false after save');

        // 저장된 content가 유효한 ZIP인지
        if (lastWrittenContent) {
            const savedZip = await JSZip.loadAsync(lastWrittenContent);
            const savedMime = await savedZip.file('mimetype')?.async('string');
            assertEqual(savedMime, 'application/hwp+zip', 'save: valid ZIP with mimetype');

            // 수정된 내용이 포함되어 있는지
            const savedSection = await savedZip.file(FILE)?.async('string');
            assert(savedSection !== undefined, 'save: section0.xml exists in saved ZIP');
            assertIncludes(savedSection || '', 'JSON으로 수정됨', 'save: saved content reflects PUT changes');
        }
    }

    // ============================================================
    // PUT text: 구조 보존 (whitespace #text 건드리지 않음, hp:linesegarray 제거)
    // ============================================================
    console.log('\n--- PUT text: structure preservation ---');
    {
        // 실제 HWPX와 유사한 들여쓰기 + linesegarray 포함 문서
        const structZip = new JSZip();
        structZip.file('mimetype', 'application/hwp+zip');
        structZip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p>
    <hp:run charPrIDRef="29">
      <hp:t>원본 텍스트</hp:t>
    </hp:run>
    <hp:linesegarray>
      <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="800" spacing="600" horzpos="0" horzsize="66432" flags="0"/>
    </hp:linesegarray>
  </hp:p>
  <hp:p>
    <hp:run charPrIDRef="10">
      <hp:t>두번째 문단</hp:t>
    </hp:run>
  </hp:p>
</hs:sec>`);
        const structUri = mockVscode.Uri.file('/test/struct.hwpx') as any;
        server.registerDocument(structUri, structZip);
        const structDoc = encodeURIComponent('/test/struct.hwpx');
        const structFile = encodeURIComponent('Contents/section0.xml');

        // hp:p 레벨에서 텍스트 수정
        const putRes = await authRequest(port, token, 'PUT',
            `/api/element?doc=${structDoc}&file=${structFile}&xpath=${encodeURIComponent('/hs:sec/hp:p[0]')}`,
            JSON.stringify({ text: '수정된 텍스트' }));
        assertEqual(putRes.status, 200, 'struct: put text on hp:p → 200');

        // hp:run의 charPrIDRef 속성이 보존되는지 확인
        const getP = await authRequest(port, token, 'GET',
            `/api/element?doc=${structDoc}&file=${structFile}&xpath=${encodeURIComponent('/hs:sec/hp:p[0]')}`);
        assertIncludes(getP.json.xml, '수정된 텍스트', 'struct: text modified');
        assertIncludes(getP.json.xml, 'charPrIDRef', 'struct: hp:run attributes preserved');

        // hp:linesegarray는 제거됨 (레이아웃 무효화)
        assert(!getP.json.xml.includes('hp:linesegarray'), 'struct: linesegarray removed');

        // 두번째 hp:p가 영향 받지 않는지 확인
        const getP2 = await authRequest(port, token, 'GET',
            `/api/element?doc=${structDoc}&file=${structFile}&xpath=${encodeURIComponent('/hs:sec/hp:p[1]')}`);
        assertIncludes(getP2.json.xml, '두번째 문단', 'struct: second paragraph preserved');
        assertIncludes(getP2.json.xml, 'charPrIDRef="10"', 'struct: second paragraph attrs preserved');

        // 원본 텍스트가 bare text로 hp:p 직하에 남지 않는지 (hp:t 내부에만 있어야 함)
        // xml 응답에서 hp:t 태그 외부에 '수정된 텍스트'가 없어야 함
        const xmlOut = getP.json.xml;
        const textOutsideHpT = xmlOut.replace(/<hp:t[^>]*>.*?<\/hp:t>/g, '').includes('수정된 텍스트');
        assert(!textOutsideHpT, 'struct: text only inside hp:t, not bare in hp:p');

        server.unregisterDocument(structUri);
    }

    // ============================================================
    // PUT text: lineBreak 보존 (다른 행의 hp:lineBreak가 roundtrip 후 유지됨)
    // ============================================================
    console.log('\n--- PUT text: lineBreak preservation across rows ---');
    {
        const lbZip = new JSZip();
        lbZip.file('mimetype', 'application/hwp+zip');
        // 표의 두 행: row1에 텍스트만, row2에 lineBreak 포함
        lbZip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:tbl>
    <hp:tr>
      <hp:tc>
        <hp:subList>
          <hp:p>
            <hp:run charPrIDRef="5">
              <hp:t>행1 텍스트</hp:t>
            </hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc>
        <hp:subList>
          <hp:p>
            <hp:run charPrIDRef="5">
              <hp:t>행2 첫줄</hp:t>
              <hp:lineBreak/>
              <hp:t>행2 둘째줄</hp:t>
            </hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
</hs:sec>`);
        const lbUri = mockVscode.Uri.file('/test/linebreak.hwpx') as any;
        server.registerDocument(lbUri, lbZip);
        const lbDoc = encodeURIComponent('/test/linebreak.hwpx');
        const lbFile = encodeURIComponent('Contents/section0.xml');

        // row1의 텍스트 수정
        const putLb = await authRequest(port, token, 'PUT',
            `/api/element?doc=${lbDoc}&file=${lbFile}&xpath=${encodeURIComponent('/hs:sec/hp:tbl/hp:tr[0]/hp:tc/hp:subList/hp:p/hp:run/hp:t')}`,
            JSON.stringify({ text: '수정된 행1' }));
        assertEqual(putLb.status, 200, 'linebreak: put text on row1 → 200');

        // row2의 XML을 확인 — lineBreak가 보존되어야 함
        const getRow2 = await authRequest(port, token, 'GET',
            `/api/element?doc=${lbDoc}&file=${lbFile}&xpath=${encodeURIComponent('/hs:sec/hp:tbl/hp:tr[1]/hp:tc/hp:subList/hp:p/hp:run')}`);
        assertEqual(getRow2.status, 200, 'linebreak: get row2 → 200');
        assertIncludes(getRow2.json.xml, '행2 첫줄', 'linebreak: row2 first text preserved');
        assertIncludes(getRow2.json.xml, '행2 둘째줄', 'linebreak: row2 second text preserved');
        assertIncludes(getRow2.json.xml, 'lineBreak', 'linebreak: hp:lineBreak preserved in row2');

        // 전체 XML도 확인 — lineBreak가 self-closing 형태로 유지
        const getFullXml = await authRequest(port, token, 'GET',
            `/api/xml?doc=${lbDoc}&file=${lbFile}`);
        assertIncludes(getFullXml.body, 'lineBreak', 'linebreak: lineBreak exists in full XML');
        assertIncludes(getFullXml.body, '수정된 행1', 'linebreak: row1 text modified in full XML');
        assertIncludes(getFullXml.body, '행2 첫줄', 'linebreak: row2 text intact in full XML');

        server.unregisterDocument(lbUri);
    }

    // ============================================================
    // Multiple documents
    // ============================================================
    console.log('\n--- Multiple documents ---');
    {
        const zip2 = new JSZip();
        zip2.file('mimetype', 'application/hwp+zip');
        zip2.file('Contents/section0.xml', '<hp:sec xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"><hp:p><hp:run><hp:t>doc2</hp:t></hp:run></hp:p></hp:sec>');
        const fakeUri2 = mockVscode.Uri.file('/test/doc2.hwpx') as any;
        server.registerDocument(fakeUri2, zip2);

        // doc param 없이 → 400
        const res = await authRequest(port, token, 'GET', '/api/files');
        assertEqual(res.status, 400, 'multi: no doc → 400');
        assert(res.json.available.length === 2, 'multi: 2 available');

        // doc param 지정
        const res2 = await authRequest(port, token, 'GET', `/api/files?doc=${encodeURIComponent('/test/doc2.hwpx')}`);
        assertEqual(res2.status, 200, 'multi: doc param → 200');

        // 존재하지 않는 doc
        const res3 = await authRequest(port, token, 'GET', `/api/files?doc=${encodeURIComponent('/test/nope.hwpx')}`);
        assertEqual(res3.status, 404, 'multi: wrong doc → 404');
        assert(Array.isArray(res3.json.available), 'multi: shows available docs');

        server.unregisterDocument(fakeUri2);
    }

    // ============================================================
    // Unregister document
    // ============================================================
    console.log('\n--- Unregister document ---');
    {
        const zip3 = new JSZip();
        zip3.file('mimetype', 'application/hwp+zip');
        const fakeUri3 = mockVscode.Uri.file('/test/temp.hwpx') as any;
        server.registerDocument(fakeUri3, zip3);

        let docsRes = await authRequest(port, token, 'GET', '/api/documents');
        const countBefore = docsRes.json.documents.length;

        server.unregisterDocument(fakeUri3);

        docsRes = await authRequest(port, token, 'GET', '/api/documents');
        assertEqual(docsRes.json.documents.length, countBefore - 1, 'unregister: document count decreased');
        const found = docsRes.json.documents.some((d: any) => d.path === '/test/temp.hwpx');
        assertEqual(found, false, 'unregister: document no longer listed');
    }

    // ============================================================
    // Unknown endpoint
    // ============================================================
    console.log('\n--- Unknown endpoint ---');
    {
        const res = await authRequest(port, token, 'GET', '/api/unknown');
        assertEqual(res.status, 404, 'unknown: status 404');
        assertIncludes(res.json.error, '/api/help', 'unknown: points to help');
    }

    // ============================================================
    // POST /api/reload
    // ============================================================
    console.log('\n--- POST /api/reload ---');
    {
        let reloadCount = 0;
        server.onReload(fakeUri as any, () => { reloadCount++; });

        const res = await authRequest(port, token, 'POST', '/api/reload');
        assertEqual(res.status, 200, 'reload: status 200');
        assert(res.json.success === true, 'reload: success');
        assert(reloadCount > 0, 'reload: listener called');
    }
    {
        // reload without token → 401
        const res = await httpRequest(port, 'POST', '/api/reload');
        assertEqual(res.status, 401, 'reload: no token → 401');
    }
    {
        // save triggers reload listener
        let saveReloadCount = 0;
        server.onReload(fakeUri as any, () => { saveReloadCount++; });

        lastWrittenUri = null;
        lastWrittenContent = null;
        const res = await authRequest(port, token, 'POST', '/api/save');
        assertEqual(res.status, 200, 'save+reload: status 200');
        assert(saveReloadCount > 0, 'save+reload: reload triggered after save');
    }

    // ============================================================
    // GET /api/help (no token required)
    // ============================================================
    console.log('\n--- GET /api/help ---');
    {
        // 토큰 없이 접근 가능
        const res = await httpRequest(port, 'GET', '/api/help');
        assertEqual(res.status, 200, 'help: status 200 without token');
        assertIncludes(res.headers['content-type'] || '', 'text/markdown', 'help: markdown content-type');
        assertIncludes(res.body, '/api/documents', 'help: documents endpoint documented');
        assertIncludes(res.body, '/api/element', 'help: element endpoint documented');
        assertIncludes(res.body, '/api/save', 'help: save endpoint documented');
        assertIncludes(res.body, '/api/reload', 'help: reload endpoint documented');
        assertIncludes(res.body, 'XPath', 'help: xpath format documented');
        assertIncludes(res.body, `127.0.0.1:${port}`, 'help: includes actual port');
        assertIncludes(res.body, '/test/doc.hwpx', 'help: lists open documents');
    }

    // ============================================================
    // Done
    // ============================================================
    server.stop();

    console.log(`\n=== Results: ${passed} passed, ${failed} failed ===`);
    if (errors.length > 0) {
        console.log('\nFailures:');
        errors.forEach(e => console.log(`  ${e}`));
    }
    process.exit(failed > 0 ? 1 : 0);
}

runTests().catch(err => {
    console.error('Test runner error:', err);
    process.exit(1);
});
