import {
  Position,
  Range,
  TextDocument,
} from "vscode-languageserver-textdocument";

import * as path from "path";
import { serverLog, documents, LogKind, Scope } from "./server";
import * as fse from "fs-extra";
import * as Encoding from "encoding-japanese";
import { fileURLToPath, pathToFileURL } from "url";
import {
  Diagnostic,
  DiagnosticSeverity,
  DocumentSymbol,
  SymbolKind,
} from "vscode-languageserver";
import { URI } from "vscode-uri";
import { ByteLengthQueuingStrategy } from "stream/web";
import { isVariableDeclaration } from "typescript";

const Accessibility = {
  public: "public",
  private: "private",
};
type Accessibility = typeof Accessibility[keyof typeof Accessibility];

const Extensions = {
  cls: "cls",
  bas: "bas",
  frm: "frm",
  vbs: "vbs",
  non: "non",
} as const;
type Extensions = keyof typeof Extensions;
// type Extensions = typeof Extensions[keyof typeof Extensions];

export type LangSymbol = {
  uri: string;
  moduleName: string;
  extension: string;
  range: Range;
  name: string;
  kind: SymbolKind;
  accessibility: Accessibility;
  functionDef: string;
  child: LangSymbol[];
  scope: string;
  parent: string;
};

// <uri, symbol[]>
export let mapUriSymbols = new Map<string, LangSymbol[]>();
let mapUriDiagnostic = new Map<string, Diagnostic[]>();

export async function getDocumentSymbol(uri: string) {
  await parseFiles(uri);

  const funcInfo = mapUriSymbols.get(uri);
  if (!funcInfo) {
    return undefined;
  }
  const symbolInformationArray = createSymbolArray(funcInfo);
  return symbolInformationArray;

  function createSymbolArray(symbols: LangSymbol[]) {
    const documentSymbols = symbols
      .filter((s) => s.kind !== SymbolKind.Variable)
      .map((symbol) => {
        const documentSymbol: DocumentSymbol = {
          name: symbol.name,
          detail: "",
          kind: symbol.kind,
          range: symbol.range,
          selectionRange: symbol.range,
        };
        if (symbol.child.length) {
          documentSymbol.children = createSymbolArray(symbol.child);
        }
        return documentSymbol;
      });
    return documentSymbols;
  }
}

export function getTokenInfo(
  uri: string,
  scope: Scope,
  name: string,
  isMethod = false
) {
  const funcName = "getFunctionInfo";
  serverLog(LogKind.DEBUG, funcName, name);

  let symbolList: LangSymbol[] | undefined;

  serverLog(LogKind.DEBUG, funcName, "search local function");
  symbolList = getSymbolList(uri, scope, name, Accessibility.private);
  if (symbolList?.length) {
    return symbolList;
  }

  serverLog(LogKind.DEBUG, funcName, "search public function");
  symbolList = getSymbolList(uri, scope, name, Accessibility.public);
  if (symbolList?.length) {
    return symbolList;
  }
}

// module
// module variable
// module function
// module function variable

// module class
// module class    variable
// module class    function
// module class    function variable

function getSymbolList(
  uri: string,
  scope: Scope,
  name: string,
  accessibility: Accessibility
) {
  const funcName = "getSymbolList";
  serverLog(
    LogKind.DEBUG,
    funcName,
    `file: ${path.basename(uri)} classScope: ${
      scope.classScope
    }  functionScope: ${
      scope.functionScope
    }  name: ${name} access: ${accessibility}`
  );

  const isVbs = uri.length > 3 && uri.slice(-3).toLowerCase() === "vbs";
  const privateUri = accessibility === Accessibility.private ? [uri] : [];
  const uris =
    accessibility === Accessibility.public ? mapUriSymbols.keys() : [];

  const symbolList: LangSymbol[] = [];

  // loop all files, but one file.
  for (const iUri of privateUri) {
    // local val
    if (scope.functionScope && !scope.classScope) {
      let moduleSymbol = mapUriSymbols
        ?.get(iUri)
        ?.find((s) => stringEq(s.name, scope.functionScope))
        ?.child.find((s) => stringEq(s.name, name));
      if (moduleSymbol) {
        serverLog(LogKind.DEBUG, funcName, "return local val.");
        return [moduleSymbol];
      }

      moduleSymbol = mapUriSymbols
        ?.get(iUri)
        ?.find((s) => stringEq(s.name, name));
      if (moduleSymbol) {
        serverLog(LogKind.DEBUG, funcName, "return local func.");
        return [moduleSymbol];
      }
    }

    //class local val
    if (scope.functionScope && scope.classScope && isVbs) {
      let moduleSymbol = mapUriSymbols
        ?.get(iUri)
        ?.find((s) => stringEq(s.name, scope.classScope))
        ?.child.find((s) => stringEq(s.name, scope.functionScope))
        ?.child.find((s) => stringEq(s.name, name));
      if (moduleSymbol) {
        serverLog(LogKind.DEBUG, funcName, "return class local val.");
        return [moduleSymbol];
      }

      moduleSymbol = mapUriSymbols
        ?.get(iUri)
        ?.find((s) => stringEq(s.name, scope.classScope))
        ?.child.find((s) => stringEq(s.name, name));
      if (moduleSymbol) {
        serverLog(LogKind.DEBUG, funcName, "return class local func.");
        return [moduleSymbol];
      }
    }

    //class val class func
    if (!scope.functionScope && scope.classScope && isVbs) {
      const moduleSymbol = mapUriSymbols
        ?.get(iUri)
        ?.find((s) => stringEq(s.name, scope.classScope))
        ?.child.find((s) => stringEq(s.name, name));
      if (moduleSymbol) {
        serverLog(LogKind.DEBUG, funcName, "return class func val.");
        return [moduleSymbol];
      }
    }

    //module val module func
    if (!scope.functionScope && !scope.classScope) {
      const moduleSymbol = mapUriSymbols
        ?.get(iUri)
        ?.find((s) => stringEq(s.name, name));
      if (moduleSymbol) {
        serverLog(
          LogKind.DEBUG,
          funcName,
          "return function or variable in a module."
        );
        return [moduleSymbol];
      }
    }
    serverLog(LogKind.DEBUG, funcName, "return no list.");
  }

  // for public
  let moduleSymbolPublic: LangSymbol[] | undefined = [];
  for (const iUri of uris) {
    serverLog(LogKind.DEBUG, funcName, iUri);

    // public
    // get topLevel symbols() ie. variable, function, class
    const topLevelSymbols = mapUriSymbols
      ?.get(iUri)
      ?.filter((s) => stringEq(s.name, name))
      ?.filter((s) => s.accessibility === accessibility);
    if (topLevelSymbols && topLevelSymbols.length) {
      serverLog(
        LogKind.DEBUG,
        funcName,
        "get topLevel symbols in public. class, function, variable"
      );
      moduleSymbolPublic = moduleSymbolPublic.concat(topLevelSymbols);
    }

    // public
    // get class symbols
    if (isVbs) {
      // do only in a class
      const symbolsInClass = mapUriSymbols
        ?.get(iUri)
        ?.filter((s) => s.kind === SymbolKind.Class);

      // get public functions and variables in class
      // local variables are open to outside
      for (const iClass of symbolsInClass ?? []) {
        const classFunctions = iClass.child.filter(
          (s) =>
            stringEq(s.name, name) && s.accessibility === Accessibility.public
        );

        if (!!classFunctions && !!classFunctions.length) {
          serverLog(
            LogKind.DEBUG,
            funcName,
            "get public function or variable in classes."
          );
          moduleSymbolPublic = moduleSymbolPublic?.concat(classFunctions);
        }
      }
    }
  }

  //
  if (moduleSymbolPublic && moduleSymbolPublic.length) {
    return moduleSymbolPublic;
  }
  serverLog(LogKind.DEBUG, funcName, "no symbols.");
}

function stringEq(s1: string, s2: string) {
  return s1.toLocaleLowerCase() === s2.toLocaleLowerCase();
}

/**
 *
 * @param uri
 */
async function parseFiles(uri: string) {
  const moduleFolder = fileURLToPath(path.dirname(uri)).toString();
  const oldMapUriSymbols = mapUriSymbols;
  mapUriSymbols = new Map<string, LangSymbol[]>();
  const oldMapUriDiagnostic = mapUriDiagnostic;
  mapUriDiagnostic = new Map<string, Diagnostic[]>();
  // get files in folder
  const files = fse.readdirSync(moduleFolder);
  const extensions = [".bas", ".frm", ".cls", ".vbs"];

  for (const file of files) {
    const fullPath = path.resolve(moduleFolder, file);
    const isFile = fse.statSync(fullPath).isFile();
    const ext = path.extname(file).toLowerCase();

    if (isFile && extensions.includes(ext)) {
      const uriAtFolder = URI.file(fullPath).toString();
      if (!oldMapUriSymbols.has(uriAtFolder) || uriAtFolder === uri) {
        const r = await parseFile(uriAtFolder);
        mapUriSymbols.set(uriAtFolder, r);
      } else {
        mapUriSymbols.set(uriAtFolder, oldMapUriSymbols.get(uriAtFolder)!);
        mapUriDiagnostic.set(
          uriAtFolder,
          oldMapUriDiagnostic.get(uriAtFolder)!
        );
      }
    }
  }
}

async function parseFile(uri: string) {
  let fileContents = documents.get(uri);
  if (!fileContents) {
    const fileFs = URI.parse(uri).fsPath.toString();
    const textFs = fse.readFileSync(fileFs);
    let docFile = Encoding.convert(textFs, {
      to: "UNICODE",
      type: "string",
    });

    docFile = docFile.replace(/\r/g, "");
    fileContents = TextDocument.create(uri, "vb", 0, docFile);
  }

  const { extensionType, moduleName } = getModuleInfo(uri);
  const rangeDummy = {
    start: { line: 0, character: 0 },
    end: { line: 0, character: 0 },
  };
  const symbol: LangSymbol = {
    uri,
    moduleName,
    extension: extensionType,
    range: rangeDummy,
    name: "",
    kind: SymbolKind.Null,
    accessibility: Accessibility.public,
    functionDef: "",
    child: [],
    scope: "",
    parent: "",
  };

  const parentTree: LangSymbol[] = [];
  parentTree.push(symbol);

  const diagnostic: Diagnostic[] = [];
  try {
    for (let line = 0; line < fileContents.lineCount; line++) {
      const currentLine = getLineAt(fileContents, line);
      const multiLine = getMultiLineAt(currentLine, fileContents, line);
      parseLine(
        fileContents.uri,
        multiLine.context,
        line,
        parentTree,
        diagnostic
      );
      line = multiLine.line;
    }
  } catch (e) {
    console.log(e);
  }
  return symbol.child;
}

function getMultiLineAt(
  currentLine: string,
  fileContents: TextDocument,
  line: number
) {
  let r = currentLine;
  let rLine = line;
  if (r.length > 0 && r.slice(-1) === "_") {
    r = r.slice(0, -1);
    for (let line2 = line + 1; line2 < fileContents.lineCount; line2++) {
      const nextLine = getLineAt(fileContents, line2);
      if (nextLine.length > 0 && nextLine.slice(-1) === "_") {
        r += nextLine.slice(0, -1);
      } else {
        r += nextLine;
        rLine = line2;
        break;
      }
    }
  }
  return { context: r, line: rLine };
}

function getLineAt(fileContents: TextDocument, line: number) {
  const r = {
    start: { line, character: 0 },
    end: { line, character: Number.MAX_VALUE },
  };
  let currentLine = fileContents.getText(r);
  currentLine = currentLine.replace(/'[^"\n]*$/gm, "").trim();
  return currentLine;
}

//
function parseLine(
  uri: string,
  contentLine: string,
  line: number,
  parentTree: LangSymbol[],
  diagnostic: Diagnostic[]
) {
  const commentReg = /'[^"\n]*$/gm;
  let currentLine = contentLine.replace(commentReg, "");
  if (contentLine !== currentLine) {
    const m = contentLine.match(commentReg);
  }
  currentLine = parseFunction(uri, currentLine, line, parentTree, diagnostic);
  currentLine = parseVariable(line, currentLine, parentTree);
  currentLine = parseEnd(uri, currentLine, line, parentTree, diagnostic);
}

function parseEnd(
  uri: string,
  contentLine: string,
  line: number,
  parentTree: LangSymbol[],
  diagnostic: Diagnostic[]
) {
  if (!contentLine) {
    return "";
  }
  const endRegex =
    /^\s*((End)\s+)((Static|Declare|Declare PtrSafe)\s+)?(Function|Sub|Class)/i;
  const endMatches = contentLine.match(endRegex);
  if (endMatches) {
    parentTree.pop();
    return "";
  }
  return contentLine;
}

function parseFunction(
  uri: string,
  contentLine: string,
  line: number,
  parentTree: LangSymbol[],
  diagnostic: Diagnostic[]
) {
  if (!contentLine) {
    return "";
  }
  const functionRegex =
    /^\s*((Public|Private|Friend)\s+)?((Static|Declare|Declare PtrSafe)\s+)?(Function|Sub|class)\s+(\w+)(\((.*)\))?/i;
  const functionMatches = contentLine.match(functionRegex);

  if (functionMatches) {
    const accessibilityString = functionMatches[2]?.toLowerCase();
    const accessibility: Accessibility =
      accessibilityString === "private"
        ? Accessibility.private
        : Accessibility.public;
    const rangeFunc: Range = {
      start: { line: line, character: 0 },
      end: { line: line, character: 0 },
    };
    const { extensionType, moduleName } = getModuleInfo(uri);
    const posDeclare = functionMatches[4]?.toString().toLowerCase() ?? "";
    const isDeclare = ["declare", "decare ptrsafe"].includes(posDeclare);
    let kind: SymbolKind = stringEq(functionMatches[5], "class")
      ? SymbolKind.Class
      : SymbolKind.Function;
    kind = isDeclare ? SymbolKind.Interface : kind;
    const symbol: LangSymbol = {
      uri,
      moduleName,
      extension: extensionType,
      range: rangeFunc,
      name: functionMatches[6].toString(),
      kind,
      accessibility,
      functionDef: contentLine.trim(),
      child: [],
      scope: "",
      parent: "",
    };

    // for global information
    parentTree[parentTree.length - 1].child.push(symbol);
    if (!isDeclare) {
      // at here scope(function, class) in
      // from here to end scope, notice the symbol to add
      parentTree.push(symbol);
      getFunctionParameter(line, functionMatches[8], parentTree);
    }

    return "";
  }
  return contentLine;
}

function parseVariable(
  line: number,
  contentLine: string,
  parentTree: LangSymbol[]
) {
  serverLog(LogKind.NONE, "getParameterOrDim Line", contentLine);

  if (!contentLine) {
    return "";
  }
  const parent = parentTree[parentTree.length - 1];

  let regex;
  let result_array;
  if (!result_array) {
    regex = /^\s*((Public|Private)\s+)?((const|dim)\s+)(.*)/i;
    result_array = contentLine.match(regex);
  }

  if (!result_array) {
    regex = /^\s*((Public|Private)\s+)((const)\s+)?(.*)/i;
    result_array = contentLine.match(regex);
  }

  if (result_array) {
    const kind =
      result_array[4] && stringEq(result_array[4], "const")
        ? SymbolKind.Constant
        : SymbolKind.Variable;
    const parameters = splitParameters(result_array[5]);
    for (let i = 0; i < parameters.length; i++) {
      //const accessibility = result_array[1] ? result_array[1] : Accessibility.public;
      const symbol: LangSymbol = {
        uri: parent.uri,
        moduleName: parent.moduleName,
        extension: parent.extension,
        range: {
          start: { line: line, character: 0 },
          end: { line: line, character: 0 },
        },
        name: parameters[i] ?? "",
        kind,
        accessibility: result_array[1] ?? Accessibility.public,
        functionDef: contentLine.trim(),
        child: [],
        scope: "",
        parent: "",
      };
      parentTree[parentTree.length - 1].child.push(symbol);
      serverLog(LogKind.NONE, "getParameterOrDim", result_array[0]);
    }
    return "";
  }
  return contentLine;
}

function getFunctionParameter(
  line: number,
  contentLine: string | undefined,
  parentTree: LangSymbol[]
) {
  if (!contentLine) {
    return;
  }
  const parent = parentTree[parentTree.length - 1];
  const kind = SymbolKind.Variable;
  const parameters = splitParameters(contentLine);
  for (let i = 0; i < parameters.length; i++) {
    const symbol: LangSymbol = {
      uri: parent.uri,
      moduleName: parent.moduleName,
      extension: parent.extension,
      range: {
        start: { line: line, character: 0 },
        end: { line: line, character: 0 },
      },
      name: parameters[i] ?? "",
      kind,
      accessibility: Accessibility.private,
      functionDef: contentLine.trim(),
      child: [],
      scope: "",
      parent: "",
    };
    parentTree[parentTree.length - 1].child.push(symbol);
  }
}

export function splitParameters(parameters: string) {
  const regex = /,(?=([^()"]*[()"][^()"]*[()"])*[^()"]*$)/;
  const parameterList = parameters.split(regex).filter(Boolean);

  const parametersReturn = parameterList
    .map((p) => {
      const regex2 =
        /(optional\s+)?((ByVal|ByRef)\s+)?(\w+)(\(.*\))?(\s+As\s\w+)*/i;
      const match = p.match(regex2);
      if (match) {
        return match[4];
      }
      return undefined;
    })
    .filter(Boolean);
  return parametersReturn;
}

// function validateTextDocument(textDocument: TextDocument) {
//   const diagnostic: Diagnostic = {
//     severity: DiagnosticSeverity.Warning,
//     range: {
//       start: textDocument.positionAt(m.index),
//       end: textDocument.positionAt(m.index + m[0].length),
//     },
//     message: `diagnostics message`,
//     source: "ex",
//   };

//   diagnostic.relatedInformation = [
//     {
//       location: {
//         uri: textDocument.uri,
//         range: Object.assign({}, diagnostic.range),
//       },
//       message: "Spelling matters",
//     },
//     {
//       location: {
//         uri: textDocument.uri,
//         range: Object.assign({}, diagnostic.range),
//       },
//       message: "Particularly for names",
//     },
//   ];
//   return diagnostic;
// }

export function getModuleInfo(uriOrPath: string) {
  const moduleFileName = path.basename(uriOrPath);
  const fileNameSplit = moduleFileName.split(".");
  // if *.sht.cls, you get *
  const moduleName = fileNameSplit[0];
  const extensionType = fileNameSplit[fileNameSplit.length - 1];
  return { extensionType, moduleName };
}
