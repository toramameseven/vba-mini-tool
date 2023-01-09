# Vba mini tool

This is the **Vba mini tool**. 

## Features

Vba mini tool supports a symbol provider and a definition provider.

This may work on files; *.cls, *.bas, *.frm and *.vbs. associated to Language[Visual Basic].


## Requirements

* Windows 10

## How to use

* Set language **Visual Basic**. Sample vscode settings.

```json
  "files.associations": {
    "*.bas": "vb",
    "*.cls": "vb",
    "*.frm": "vb",
    "*.vbs": "vb",
  },
```

* Hover a symbol, shows definition.
* Goto definition or Ctrl + Left click, goes to the definition.

## Extension Settings

* VbaMiniTool.outline.variable

  when true, shows variables in the outline
* VbaMiniTool.outline.constant

  when true, shows constants in the outline
* VbaMiniTool.outline.declare

  when true, shows declare functions in the outline

After changing the configurations, the outline is not refreshed.
Edit the file or reopen the file.


## Known Issues

* This extension works for only the modules in the same folder.
* Since the context is not taken into account, this extension detects all same symbols.
* When you jump to the definition, the position may be near near the target position.
* Sometimes, multi lines(with '_') works fail.
* Properties are shown as *Function* in the outline.
* Properties works so so. Sorry.
* The features we do not recognize do not work well.
* Api declares are shown as interface.
* Vba interface is not supported.

## How to package and publish

1. npm install -g vsce
2. vsce package --target win32-x64
3. vsce publish

## Acknowledgments

We thank you for the wonderful npm packages.

[VSCode VBA](https://marketplace.visualstudio.com/items?itemName=spences10.VBA) is very useful for developing this extension. We appreciate it very much.

## Release Notes

0.0.2

fixed

* configuration is not work.

0.0.1 first release.

