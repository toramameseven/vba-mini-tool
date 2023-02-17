/* --------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License. See License.txt in the project root for license information.
 * ------------------------------------------------------------------------------------------ */

// https://github.com/microsoft/vscode-markdown-languageservice/blob/main/example.cjs

import {
  createConnection,
  TextDocuments,
  Diagnostic,
  DiagnosticSeverity,
  ProposedFeatures,
  InitializeParams,
  DidChangeConfigurationNotification,
  CompletionItem,
  CompletionItemKind,
  TextDocumentPositionParams,
  TextDocumentSyncKind,
  InitializeResult,
  HoverParams,
  Hover,
  DefinitionParams,
  Definition,
  Location,
  Range,
  DocumentSymbolParams,
  SymbolInformation,
  WorkspaceSymbolParams,
  TextDocumentSyncOptions,
  Position,
  integer,
  MarkedString,
} from "vscode-languageserver/node";
import { TextDocument } from "vscode-languageserver-textdocument";
import * as VbaParser from "./vbaMiniParser";
import path = require("path");

// Create a connection for the server, using Node's IPC as a transport.
// Also include all preview / proposed LSP features.
export const connection = createConnection(ProposedFeatures.all);

// Create a simple text document manager.
export const documents: TextDocuments<TextDocument> = new TextDocuments(
  TextDocument
);

export type Scope = { classScope: string; functionScope: string };

let hasConfigurationCapability = false;
let hasWorkspaceFolderCapability = false;
let hasDiagnosticRelatedInformationCapability = false;

//
connection.onInitialize((params: InitializeParams) => {
  serverLog(LogKind.TRACE, "onInitialize");
  const capabilities = params.capabilities;

  // Does the client support the `workspace/configuration` request?
  // If not, we fall back using global settings.
  hasConfigurationCapability = !!(
    capabilities.workspace && !!capabilities.workspace.configuration
  );
  hasWorkspaceFolderCapability = !!(
    capabilities.workspace && !!capabilities.workspace.workspaceFolders
  );
  hasDiagnosticRelatedInformationCapability = !!(
    capabilities.textDocument &&
    capabilities.textDocument.publishDiagnostics &&
    capabilities.textDocument.publishDiagnostics.relatedInformation
  );
  const syncOptions: TextDocumentSyncOptions = {
    openClose: true,
    change: TextDocumentSyncKind.Full,
  };

  const result: InitializeResult = {
    capabilities: {
      textDocumentSync: syncOptions,
      hoverProvider: true,
      definitionProvider: true,
      documentSymbolProvider: true,
    },
  };

  if (hasWorkspaceFolderCapability) {
    result.capabilities.workspace = {
      workspaceFolders: {
        supported: true,
      },
    };
  }
  return result;
});

//
connection.onInitialized(async () => {
  if (hasConfigurationCapability) {
    // Register for all configuration changes.
    connection.client.register(
      DidChangeConfigurationNotification.type,
      undefined
    );
  }
  if (hasWorkspaceFolderCapability) {
    connection.workspace.onDidChangeWorkspaceFolders((_event) => {
      serverLog(LogKind.TRACE, "Workspace folder change event received.");
    });
  }

  globalSettings = await connection.workspace.getConfiguration(
    "VbaMiniTool.outline"
  );

  serverLog(LogKind.INFO, "onInitialized OK!!");
});

//
connection.onDocumentSymbol((params: DocumentSymbolParams) => {
  // uri file:///c%3A/projects/hub-toramame/lsp-sample-vba/vbaSample/test.bas
  serverLog(LogKind.TRACE, "onDocumentSymbol", params.textDocument.uri);
  const symbols = VbaParser.getDocumentSymbol(
    params.textDocument.uri,
    globalSettings
  );
  return symbols;
});

//
connection.onDefinition((params: DefinitionParams) => {
  const funcName = "connection.onDefinition";
  serverLog(LogKind.TRACE, funcName, params.textDocument.uri);
  const { word, isAlone } = getWordAtPosition(
    params.textDocument.uri,
    params.position
  );
  const scopes = getScope(params.textDocument.uri, params.position);
  const funcInfos = VbaParser.getSymbolList(
    params.textDocument.uri,
    scopes,
    word,
    isAlone
  );

  if (!funcInfos || funcInfos?.length === 0) {
    serverLog(LogKind.TRACE, "definition undefined!!");
    return undefined;
  }

  const definitionContent = funcInfos[0].range.start;
  if (!definitionContent) {
    serverLog(LogKind.TRACE, "definition undefined!!");
    return undefined;
  }

  const definition: Location[] = funcInfos.map((s) => {
    serverLog(LogKind.DEBUG, funcName, "uri", s.uri);
    return { uri: s.uri, range: s.range };
  });
  return definition;
});

export interface VbaMiniSettings {
  variable: boolean;
  constant: boolean;
  declare: boolean;
}

const defaultSettings: VbaMiniSettings = {
  variable: true,
  constant: true,
  declare: true,
};

let globalSettings: VbaMiniSettings = defaultSettings;

// now not use.
connection.onDidChangeConfiguration(async (change) => {
  globalSettings = await connection.workspace.getConfiguration(
    "VbaMiniTool.outline"
  );

  serverLog(
    LogKind.TRACE,
    `${globalSettings.constant}  ${globalSettings.variable}  ${globalSettings.declare}`
  );
  serverLog(LogKind.TRACE, "onDidChangeConfiguration");
});

//
connection.onHover((params: HoverParams) => {
  const funcName = "connection.onHover";
  serverLog(LogKind.DEBUG, funcName, params.textDocument.uri);

  const { word, isAlone } = getWordAtPosition(
    params.textDocument.uri,
    params.position
  );

  const scopes = getScope(params.textDocument.uri, params.position); //

  serverLog(LogKind.DEBUG, funcName, scopes.classScope, scopes.functionScope);

  const funcInfos = VbaParser.getSymbolList(
    params.textDocument.uri,
    scopes,
    word,
    // function: f()
    // method: a.f()
    isAlone
  );

  if (!funcInfos || funcInfos?.length === 0) {
    serverLog(LogKind.DEBUG, "hover no fun infos !!");
    return undefined;
  }

  const hoverContent = funcInfos[0].range.start;
  if (!hoverContent) {
    serverLog(LogKind.DEBUG, "hover undefined!!");
    return undefined;
  }

  const defs = funcInfos.length;
  // if detects multi defs, add moduleName.
  const functionDefs = funcInfos
    .map((f) => f.functionDef + (defs > 1 ? ": " + path.basename(f.uri) : ""))
    .join("\n");

  const hover: Hover = {
    contents: { language: "vb", value: functionDefs },
  };
  serverLog(LogKind.TRACE, funcName, `${scopes} : ${word}`);
  return hover;
});

//
function getScope(uri: string, position: Position): Scope {
  const funcName = "getScope";
  serverLog(LogKind.TRACE, funcName, "in", uri);
  const document = documents.get(uri);
  let classScope = "";
  let functionScope = "";
  if (!document) {
    return { classScope, functionScope };
  }

  const isVbs = uri.length > 3 && uri.slice(-3).toLowerCase() === "vbs";

  // 8, (Function|Sub|Class)
  // 7, (end)
  // 6, (\w+) function name
  // 5, (Function|Sub|class)
  const functionRegex =
    /^\s*((Public|Private|Friend)\s+)?((Static|Declare|Declare PtrSafe)\s+)?(Function|Sub|class|Property get|Property set|Property let)\s+(\w+)|^\s*(End)\s+(Function|Sub|Class)/i;
  for (let testLine = position.line; testLine > -1; testLine--) {
    const line = document.getText({
      start: { line: testLine, character: 0 },
      end: { line: testLine, character: integer.MAX_VALUE },
    });

    const matches = line.match(functionRegex);
    if (matches) {
      if (
        matches[7]?.toLowerCase() === "end" &&
        matches[8]?.toLowerCase() === "class"
      ) {
        serverLog(LogKind.TRACE, funcName, "match a class End");
        classScope = ""; //matches[8].toLowerCase();
        // outside vbs class
        return { classScope, functionScope };
        // } else if (!functionScope && matches[7]?.toLowerCase() === "end") {
        //   // end function or end sub
        //   functionScope = matches[8];
        //   // next class search
        //   serverLog(
        //     LogKind.TRACE,
        //     funcName,
        //     "match to sub or function",
        //     matches[6]
        //   );
      } else if (!classScope && matches[5]?.toLowerCase() === "class") {
        classScope = matches[6];
        return { classScope, functionScope };
      } else if (!functionScope) {
        functionScope = matches[6];
      }
    }
    if ((!!classScope && !!functionScope) || (!!functionScope && !isVbs)) {
      break;
    }
  }
  return { classScope, functionScope };
}

//
function getWordAtPosition(uri: string, position: Position) {
  const character = position.character;
  const document = documents.get(uri);
  const line = document?.getText({
    start: { line: position.line, character: 0 },
    end: { line: position.line, character: integer.MAX_VALUE },
  });

  if (!line) {
    return { word: "", isAlone: false };
  }

  const functionRegex = /\.?\w+/g;
  let lastMatch = "";
  let lastStart = 0;
  let lastEnd = 0;

  for (const m of (line + " dummy").matchAll(functionRegex)) {
    const matchIndex = m.index ?? 0;
    if (lastStart <= character && character < lastEnd) {
      const isMethod = lastMatch.length > 0 && lastMatch[0] === ".";
      const word = isMethod ? lastMatch.slice(1) : lastMatch;
      return { word, isAlone: !isMethod };
    }
    lastMatch = m[0];
    lastStart = matchIndex;
    lastEnd = matchIndex + lastMatch.length;
  }
  return { word: "", isAlone: false };
}

// The content of a text document has changed. This event is emitted
// when the text document first opened or when its content has changed.
documents.onDidChangeContent((change) => {
  serverLog(LogKind.TRACE, `onDidChangeContent: ${change.document.uri}`);
});

//
connection.onDidChangeTextDocument((handler) => {
  serverLog(
    LogKind.TRACE,
    `onDidChangeTextDocument: ${handler.textDocument.uri}`
  );
});

//
connection.onDidChangeWatchedFiles((_change) => {
  // Monitored files have change in VSCode
  serverLog(LogKind.TRACE, "onDidChangeWatchedFiles");
});

// Make the text document manager listen on the connection
// for open, change and close text document events
documents.listen(connection);

// Listen on the connection
connection.listen();

export const LogKind = {
  NONE: 0,
  TRACE: 1,
  DEBUG: 2,
  INFO: 3,
  WARN: 4,
  ERROR: 5,
  FATAL: 6,
} as const;
export type LogKind = typeof LogKind[keyof typeof LogKind];

//
export function serverLog(
  logKind: LogKind,
  title: string,
  message = "",
  message2 = ""
) {
  if (
    [
      //LogKind.TRACE as number,
      LogKind.INFO as number,
      LogKind.ERROR as number,
      //LogKind.DEBUG as number,
    ].includes(logKind)
  ) {
    connection.console.log(
      `====> : ${new Date().toLocaleTimeString()}  : ${
        Object.keys(LogKind)[logKind]
      }: ${title}: ${message}: ${message2}`
    );
  }
}
