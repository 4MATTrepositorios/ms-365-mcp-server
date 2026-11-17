import { describe, it, expect, vi } from 'vitest';
import { UTILITY_TOOLS, UPLOAD_FILE_MAX_BYTES, uploadFileEndpoint } from '../src/graph-tools.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn(), verbose: vi.fn() },
  enableConsoleLogging: vi.fn(),
}));

const tool = UTILITY_TOOLS.find((t) => t.name === 'upload-file-to-folder')!;
const FOLDER = '/sites/contoso.sharepoint.com,abc,def/drive/root';

type Call = { endpoint: string; options: Record<string, unknown> };

function ctx(reply: unknown = { id: 'NEW', name: 'ata.md', webUrl: 'https://x/ata.md', size: 5 }) {
  const calls: Call[] = [];
  const graphClient = {
    graphRequest: vi.fn(async (endpoint: string, options: Record<string, unknown>) => {
      calls.push({ endpoint, options });
      if (reply instanceof Error) {
        return {
          content: [{ type: 'text', text: JSON.stringify({ error: reply.message }) }],
          isError: true,
          _meta: { http_status: 409 },
        };
      }
      return {
        content: [{ type: 'text', text: JSON.stringify(reply) }],
        _meta: { http_status: 201 },
      };
    }),
  };
  return {
    calls,
    ctx: { graphClient: graphClient as never, multiAccount: false, accountNames: [] },
  };
}

function parse(result: { content: Array<{ type: string; text?: string }> }) {
  return JSON.parse(result.content[0].text as string);
}

describe('upload-file-to-folder', () => {
  it('is a write tool that never offers replace', () => {
    expect(tool).toBeDefined();
    expect(tool.readOnlyHint).toBe(false);
    const schema = tool.buildSchema({ multiAccount: false, accountNames: [] } as never);
    const behaviors = (
      schema.conflictBehavior as unknown as { unwrap: () => { options: string[] } }
    ).unwrap().options;
    expect(behaviors).toEqual(['rename', 'fail']);
  });

  it('builds the path-based create URL with an encoded name and the pinned behavior', () => {
    expect(uploadFileEndpoint(FOLDER, 'Ata de reunião 1.md', 'rename')).toBe(
      `${FOLDER}:/Ata%20de%20reuni%C3%A3o%201.md:/content?@microsoft.graph.conflictBehavior=rename`
    );
  });

  it('PUTs UTF-8 text with rename by default and reports the stored file', async () => {
    const { calls, ctx: c } = ctx();
    const out = await tool.execute({ folder: FOLDER, name: 'ata.md', content: 'olá' }, c);
    expect(out.isError).toBeFalsy();
    expect(calls).toHaveLength(1);
    expect(calls[0].endpoint).toBe(
      `${FOLDER}:/ata.md:/content?@microsoft.graph.conflictBehavior=rename`
    );
    expect(calls[0].options.method).toBe('PUT');
    expect(Buffer.from(calls[0].options.body as Buffer).toString('utf8')).toBe('olá');
    expect((calls[0].options.headers as Record<string, string>)['Content-Type']).toBe(
      'text/plain; charset=utf-8'
    );
    expect(parse(out as never)).toEqual({
      id: 'NEW',
      name: 'ata.md',
      webUrl: 'https://x/ata.md',
      size: 5,
      renamed: false,
    });
  });

  it('says when Graph stored the file under another name', async () => {
    const { ctx: c } = ctx({ id: 'N2', name: 'ata 1.md', webUrl: 'u', size: 1 });
    const out = await tool.execute({ folder: FOLDER, name: 'ata.md', content: 'x' }, c);
    expect(parse(out as never).renamed).toBe(true);
  });

  it('decodes base64 content into bytes', async () => {
    const { calls, ctx: c } = ctx();
    await tool.execute(
      {
        folder: '/drives/b!xyz/items/01ABC',
        name: 'a.bin',
        content: 'AAEC/w==',
        encoding: 'base64',
      },
      c
    );
    expect([...Buffer.from(calls[0].options.body as Buffer)]).toEqual([0, 1, 2, 255]);
    expect((calls[0].options.headers as Record<string, string>)['Content-Type']).toBe(
      'application/octet-stream'
    );
  });

  it('refuses replace, bad folders, path-like names, bad base64 and oversize content without calling Graph', async () => {
    const { calls, ctx: c } = ctx();
    const bad = [
      { folder: FOLDER, name: 'a.md', content: 'x', conflictBehavior: 'replace' },
      { folder: '/sites/s1/lists/l1', name: 'a.md', content: 'x' },
      { folder: '/me/messages/1/attachments', name: 'a.md', content: 'x' },
      { folder: FOLDER, name: '../a.md', content: 'x' },
      { folder: FOLDER, name: 'a:b.md', content: 'x' },
      { folder: FOLDER, name: ' a.md', content: 'x' },
      { folder: FOLDER, name: 'a.md', content: 'not base64!', encoding: 'base64' },
      { folder: FOLDER, name: 'a.md', content: 'x'.repeat(UPLOAD_FILE_MAX_BYTES + 1) },
      { folder: FOLDER, name: 'a.md' },
    ];
    for (const params of bad) {
      const out = await tool.execute(params as never, c);
      expect(out.isError, JSON.stringify(params).slice(0, 80)).toBe(true);
    }
    expect(calls).toHaveLength(0);
  });

  it('explains a name conflict under fail without writing', async () => {
    const { ctx: c } = ctx(new Error('nameAlreadyExists: The specified item name already exists'));
    const out = await tool.execute(
      { folder: FOLDER, name: 'a.md', content: 'x', conflictBehavior: 'fail' },
      c
    );
    expect(out.isError).toBe(true);
    expect(parse(out as never).error).toMatch(/already exists.*Nothing was written/);
  });
});
