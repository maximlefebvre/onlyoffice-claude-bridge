// Tools available for every editor type.

import { z } from 'zod';
import { documentArg } from './shared.mjs';

export const commonTools = [
  {
    name: 'list_documents',
    title: 'List connected documents',
    description: 'List the OnlyOffice documents currently connected through the Claude Bridge plugin (id, title, editor type: word = text document, cell = spreadsheet, slide = presentation).',
    schema: {},
    readOnly: true,
    noDocument: true,
    run: (b) => b.listDocuments(),
  },

  {
    name: 'list_connection_requests',
    title: 'List connection requests',
    description: 'List plugins waiting for the user\'s approval to connect (used when no approval dialog can be shown on this computer). Always ask the user before approving.',
    schema: {},
    readOnly: true,
    noDocument: true,
    run: (b) => b.listPairingRequests(),
  },

  {
    name: 'approve_connection',
    title: 'Approve or deny a connection request',
    description: 'Approve or deny a pending plugin connection request. Only call this after the user explicitly said yes (or no) for this request in the conversation.',
    schema: {
      id: z.number().int().describe('Request id from list_connection_requests.'),
      allow: z.boolean(),
    },
    noDocument: true,
    run: (b, _c, a) => {
      if (!b.decidePairing(a.id, a.allow)) throw new Error(`No pending connection request #${a.id}`);
      return { id: a.id, allowed: a.allow };
    },
  },

  {
    name: 'manage_paired_browsers',
    title: 'List or revoke paired browsers',
    description: 'List the browsers allowed to connect (each approval pairs one browser profile), or revoke one so it must be approved again.',
    schema: {
      revoke: z.number().int().optional().describe('Id of the pairing to revoke.'),
    },
    noDocument: true,
    destructive: true,
    run: (b, _c, a) => {
      if (a.revoke !== undefined) {
        const removed = b.devices.remove(a.revoke);
        if (!removed) throw new Error(`No pairing #${a.revoke}`);
        return { revoked: a.revoke };
      }
      return b.devices.list();
    },
  },

  {
    name: 'get_selection',
    title: 'Get selected text',
    description: 'Return the text the user has currently selected in a document, spreadsheet or presentation.',
    schema: { ...documentArg },
    readOnly: true,
    run: async (b, c) => {
      const text = await b.method(c, 'GetSelectedText', [{
        Numbering: false, Math: false, TableCellSeparator: '\t', TableRowSeparator: '\n', ParaSeparator: '\n', TabSymbol: '\t',
      }]);
      return { text: text || '' };
    },
  },

  {
    name: 'replace_selection',
    title: 'Replace selected text',
    description: 'Replace the current selection with new text (or insert at the cursor when nothing is selected). Works in text documents and in presentation text boxes. Lines become paragraphs.',
    schema: { ...documentArg, text: z.string() },
    run: async (b, c, a) => {
      await b.method(c, 'PasteText', [a.text]);
      return { replaced: true };
    },
  },

  {
    name: 'run_script',
    title: 'Run an OnlyOffice API script',
    description: [
      'Advanced: run custom JavaScript inside the editor with the OnlyOffice Document Builder API, for anything the other tools do not cover.',
      '`code` is the body of a function. The global `Api` is the editor API: Api.GetDocument() (text), Api.GetActiveSheet()/Api.GetSheet(name) (spreadsheet), Api.GetPresentation() (presentation).',
      'Inputs passed in `scope` are available as `Asc.scope`. `return` a JSON-serializable value; thrown errors are reported back.',
      'Reference: https://api.onlyoffice.com/docs/office-api/get-started/overview/',
    ].join(' '),
    schema: {
      ...documentArg,
      code: z.string().min(1).describe('Function body, e.g. "var d = Api.GetDocument(); return d.GetElementsCount();"'),
      scope: z.record(z.string(), z.any()).optional(),
    },
    destructive: true,
    run: (b, c, a) => b.run(c, a.code, a.scope || {}),
  },

  {
    name: 'execute_method',
    title: 'Call an editor method',
    description: 'Advanced: call an OnlyOffice plugin editor method (window.Asc.plugin.executeMethod), e.g. "GetSelectedText", "PasteText", "SearchAndReplace", "GetAllComments". Reference: https://api.onlyoffice.com/docs/plugin-and-macros/interacting-with-editors/overview/',
    schema: {
      ...documentArg,
      name: z.string().min(1),
      args: z.array(z.any()).optional(),
    },
    destructive: true,
    run: (b, c, a) => b.method(c, a.name, a.args || []),
  },
];
