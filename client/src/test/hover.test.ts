/* --------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License. See License.txt in the project root for license information.
 * ------------------------------------------------------------------------------------------ */
//C:\home\tora-hub\vba-mini-tool\.vscode-test\user-data\User\settings.json

import * as vscode from "vscode";
import * as assert from "assert";
import { getDocUri, activate } from "./helper";

suite("suite hover test vbsSample.vbs", async () => {
  const docUri = getDocUri("vbsSample.vbs");
  setup(createExpectedVbsSample);

  test("hover test no. 1a", async () => {
    for await (const i of vbsSampleVbs) {
      await testHoversEx(docUri, i.line, i.column, i.value);
    }
  });
});

suite("=====suite hover test Class1.cls", async () => {
  const docUri = getDocUri("Class1.cls");
  setup(createExpectedClass1);

  test("hover test Class1.cls", async () => {
    for await (const i of class1Sample) {
      await testHoversEx(docUri, i.line, i.column, i.value);
    }
  });
});

suite("====suite hover test Module1", async () => {
  const docUri = getDocUri("Module1.bas");
  setup(createExpectedModule1);

  test("hover test Module1", async () => {
    for await (const i of module1Sample) {
      await testHoversEx(docUri, i.line, i.column, i.value);
    }
  });
});

//
async function createDocumentSymbols(docUri: vscode.Uri) {
  const symbols: vscode.DocumentSymbol[] = await vscode.commands.executeCommand(
    "vscode.executeDocumentSymbolProvider",
    docUri
  );
}

//
async function createSampleHovers(
  docUri: vscode.Uri,
  line: number,
  column: number
) {
  //await activate(docUri);
  const position = new vscode.Position(line - 1, column - 1);
  const actualHovers: vscode.Hover[] = await vscode.commands.executeCommand(
    "vscode.executeHoverProvider",
    docUri,
    position //new vscode.Position(14, 10)
  );
  if (actualHovers) {
    actualHovers.forEach((hover, i) => {
      hover.contents.forEach((markdownString: vscode.MarkdownString) => {
        console.log(`{
  line: ${line},
  column: ${column},
  value: "${markdownString.value.replace(/\r?\n/g, "")}"
},`);
      });
    });
  }
}

//
async function createExpectedModule1() {
  const docUri = getDocUri("Module1.bas");
  await activate(docUri);

  await createSampleHovers(docUri, 4, 12);
  await createSampleHovers(docUri, 5, 12);
  await createSampleHovers(docUri, 6, 12);
  await createSampleHovers(docUri, 9, 15);
  await createSampleHovers(docUri, 11, 15);
  await createSampleHovers(docUri, 14, 15);
  await createSampleHovers(docUri, 18, 13);
  await createSampleHovers(docUri, 19, 18);
}

//
async function createExpectedClass1() {
  const docUri = getDocUri("Class1.cls");

  await activate(docUri);
  //await createDocumentSymbols(docUri);

  await createSampleHovers(docUri, 12, 15);
  await createSampleHovers(docUri, 13, 11);
  await createSampleHovers(docUri, 14, 11);
  await createSampleHovers(docUri, 17, 24);
  await createSampleHovers(docUri, 18, 12);
  await createSampleHovers(docUri, 21, 16);
  await createSampleHovers(docUri, 22, 10);
  await createSampleHovers(docUri, 25, 28);
  await createSampleHovers(docUri, 26, 16);
  await createSampleHovers(docUri, 27, 7);
  await createSampleHovers(docUri, 12, 15);
}

//
async function createExpectedVbsSample() {
  const docUri = getDocUri("vbsSample.vbs");
  await activate(docUri);

  await createSampleHovers(docUri, 1, 9);

  await createSampleHovers(docUri, 3, 12);
  await createSampleHovers(docUri, 5, 14);
  await createSampleHovers(docUri, 6, 16);
  await createSampleHovers(docUri, 8, 25);
  await createSampleHovers(docUri, 9, 19);
  await createSampleHovers(docUri, 10, 13);
  await createSampleHovers(docUri, 16, 11);

  await createSampleHovers(docUri, 23, 13);
  await createSampleHovers(docUri, 25, 14);
  await createSampleHovers(docUri, 28, 24);
  await createSampleHovers(docUri, 29, 23);
  await createSampleHovers(docUri, 30, 14);
  await createSampleHovers(docUri, 33, 24);
  await createSampleHovers(docUri, 36, 10);

  await createSampleHovers(docUri, 41, 17);
  await createSampleHovers(docUri, 43, 16);
  await createSampleHovers(docUri, 44, 15);
  await createSampleHovers(docUri, 44, 24);
  await createSampleHovers(docUri, 46, 11);
  await createSampleHovers(docUri, 46, 25);
  await createSampleHovers(docUri, 50, 15);
  await createSampleHovers(docUri, 51, 9);
  await createSampleHovers(docUri, 52, 5);
  await createSampleHovers(docUri, 52, 15);
  await createSampleHovers(docUri, 53, 18);
  await createSampleHovers(docUri, 55, 9);
  await createSampleHovers(docUri, 56, 9);
  await createSampleHovers(docUri, 56, 24);
  await createSampleHovers(docUri, 57, 5);
  await createSampleHovers(docUri, 57, 13);
  await createSampleHovers(docUri, 58, 10);
  await createSampleHovers(docUri, 58, 14);
  await createSampleHovers(docUri, 60, 10);
  await createSampleHovers(docUri, 61, 10);
  await createSampleHovers(docUri, 61, 26);
  await createSampleHovers(docUri, 62, 6);
  await createSampleHovers(docUri, 62, 11);
  await createSampleHovers(docUri, 63, 11);
  await createSampleHovers(docUri, 63, 15);
}

async function testHoversEx(
  docUri: vscode.Uri,
  line: number,
  column: number,
  value: string
) {
  //await activate(docUri);
  const position = new vscode.Position(line - 1, column - 1);
  const actualHovers: vscode.Hover[] = await vscode.commands.executeCommand(
    "vscode.executeHoverProvider",
    docUri,
    position //new vscode.Position(14, 10)
  );

  assert.equal(actualHovers.length, 1, "no hover");
  if (actualHovers) {
    actualHovers.forEach((hover, i) => {
      hover.contents.forEach((markdownString: vscode.MarkdownString) => {
        assert.equal(
          markdownString.value.replace(/\r?\n/g, ""),
          value.replace(/\r?\n/g, "")
        );
      });
    });
  }
}

const module1Sample = [
  {
    line: 4,
    column: 12,
    value: "```vbDim Module1_Dim_Val```",
  },
  {
    line: 5,
    column: 12,
    value: "```vbPublic Module1_Public_Val```",
  },
  {
    line: 6,
    column: 12,
    value: "```vbPrivate Module1_Private_Val```",
  },
  {
    line: 9,
    column: 15,
    value: "```vbPublic Sub Module1()```",
  },
  {
    line: 11,
    column: 15,
    value: "```vbPublic Module2_Public_Val```",
  },
  {
    line: 14,
    column: 15,
    value: "```vbPublic Sub Common()```",
  },
  {
    line: 18,
    column: 13,
    value: "```vbPublic Sub Log(ByRef msg As String)```",
  },
  {
    line: 19,
    column: 18,
    value: "```vbByRef msg As String```",
  },
];

const class1Sample = [
  {
    line: 12,
    column: 15,
    value: "```vbPrivate class_private As String```",
  },
  {
    line: 13,
    column: 11,
    value: "```vbPublic class_public As String```",
  },
  {
    line: 14,
    column: 11,
    value: "```vbDim class_dim As String```",
  },
  {
    line: 17,
    column: 24,
    value: "```vbPublic Function Public_Function() As String```",
  },
  {
    line: 18,
    column: 12,
    value: "```vbPublic Function Public_Function() As String```",
  },
  {
    line: 21,
    column: 16,
    value: "```vbFunction Non_Function() As String```",
  },
  {
    line: 22,
    column: 10,
    value: "```vbFunction Non_Function() As String```",
  },
  {
    line: 25,
    column: 28,
    value: "```vbPrivate Function Private_Function() As String```",
  },
  {
    line: 26,
    column: 16,
    value: "```vbPrivate Function Private_Function() As String```",
  },
  {
    line: 27,
    column: 7,
    value:
      "```vbPublic Sub Common(): Module1.basPublic Sub Common(): Module2.bas```",
  },
  {
    line: 12,
    column: 15,
    value: "```vbPrivate class_private As String```",
  },
];

const vbsSampleVbs = [
  {
    line: 1,
    column: 9,
    value: "```vbdim vbsSample```",
  },
  {
    line: 3,
    column: 12,
    value: "```vbClass SampleClass1```",
  },
  {
    line: 5,
    column: 14,
    value: "```vbSampleClass1::dim class1_val```",
  },
  {
    line: 6,
    column: 16,
    value: "```vbSampleClass1::Public common_val```",
  },
  {
    line: 8,
    column: 25,
    value: "```vbSampleClass1::function SampleClass1_Func1()```",
  },
  {
    line: 9,
    column: 19,
    value: "```vbdim class1_Func1_val```",
  },
  {
    line: 10,
    column: 13,
    value: "```vbSampleClass1::dim class1_val```",
  },
  {
    line: 16,
    column: 11,
    value: "```vbSampleClass1::sub log()```",
  },
  {
    line: 23,
    column: 13,
    value: "```vbClass SampleClass2```",
  },
  {
    line: 25,
    column: 14,
    value: "```vbSampleClass2::dim class2_val```",
  },
  {
    line: 28,
    column: 24,
    value: "```vbSampleClass2::function SampleClass2_Func1()```",
  },
  {
    line: 29,
    column: 23,
    value: "```vbdim class2_Func1_val```",
  },
  {
    line: 30,
    column: 14,
    value: "```vbSampleClass2::dim class2_val```",
  },
  {
    line: 33,
    column: 24,
    value: "```vbSampleClass2::public function common()```",
  },
  {
    line: 36,
    column: 10,
    value: "```vbSampleClass2::sub log()```",
  },
  {
    line: 41,
    column: 17,
    value: "```vbFunction myFunction(val1)```",
  },
  {
    line: 43,
    column: 16,
    value: "```vbDim myFunction_Val```",
  },
  {
    line: 44,
    column: 15,
    value: "```vbDim myFunction_Val```",
  },
  {
    line: 44,
    column: 24,
    value: "```vbval1```",
  },
  {
    line: 46,
    column: 11,
    value: "```vbFunction myFunction(val1)```",
  },
  {
    line: 46,
    column: 25,
    value: "```vbDim myFunction_Val```",
  },
  {
    line: 50,
    column: 15,
    value: "```vbprivate Sub Main()```",
  },
  {
    line: 51,
    column: 9,
    value: "```vbDim r```",
  },
  {
    line: 52,
    column: 5,
    value: "```vbDim r```",
  },
  {
    line: 52,
    column: 15,
    value: "```vbFunction myFunction(val1)```",
  },
  {
    line: 53,
    column: 18,
    value: "```vbDim r```",
  },
  {
    line: 55,
    column: 9,
    value: "```vbDim c```",
  },
  {
    line: 56,
    column: 9,
    value: "```vbDim c```",
  },
  {
    line: 56,
    column: 24,
    value: "```vbClass SampleClass1```",
  },
  {
    line: 57,
    column: 5,
    value: "```vbDim c```",
  },
  {
    line: 57,
    column: 13,
    value: "```vbSampleClass1::Public common_val```",
  },
  {
    line: 58,
    column: 10,
    value: "```vbDim c```",
  },
  {
    line: 58,
    column: 14,
    value:
      "```vbSampleClass1::sub log(): vbsSample.vbsSampleClass2::sub log(): vbsSample.vbs```",
  },
  {
    line: 60,
    column: 10,
    value: "```vbDim c2```",
  },
  {
    line: 61,
    column: 10,
    value: "```vbDim c2```",
  },
  {
    line: 61,
    column: 26,
    value: "```vbClass SampleClass2```",
  },
  {
    line: 62,
    column: 6,
    value: "```vbDim c2```",
  },
  {
    line: 62,
    column: 11,
    value: "```vbSampleClass2::public function common()```",
  },
  {
    line: 63,
    column: 11,
    value: "```vbDim c2```",
  },
  {
    line: 63,
    column: 15,
    value:
      "```vbSampleClass1::sub log(): vbsSample.vbsSampleClass2::sub log(): vbsSample.vbs```",
  },
];
