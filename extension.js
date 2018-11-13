const vscode = require('vscode');
const path = require('path');
const classGen = require('./src/class-generators');
const moduleGen = require('./src/module-generators');

function activate(context) {
    let disposableGS = vscode.commands.registerCommand('extension.generateVBGetterAndSetters', function () {
        var editor = vscode.window.activeTextEditor;
        if (!editor)
            return; // No open text editor

        var selection = editor.selection;
        var text = editor.document.getText(selection);

        if (text.length < 1) {
            vscode.window.showErrorMessage('No selected properties.');
            return;
        }

        try {
            var getterAndSetterObj = classGen.createGetterAndSetter(text);

            if (!getterAndSetterObj.errorFlag) {
                editor.edit(
                    edit => editor.selections.forEach(
                      selection => 
                      {
                        edit.insert(selection.end, getterAndSetterObj.generatedCode);
                      }
                    )
                );
            }
        } 
        catch (error) {
            console.log(error);
            vscode.window.showErrorMessage('Something went wrong! Try that the properties are in this format: "Private p_**** As String"');
        }
    });

    let disposableConstructors = vscode.commands.registerCommand('extension.generateVBConstructor', function () {
        var editor = vscode.window.activeTextEditor;
        if (!editor)
            return; // No open text editor

        var selection = editor.selection;
        var text = editor.document.getText(selection);

        if (text.length < 1) {
            vscode.window.showErrorMessage('No selected properties.');
            return;
        }

        try {
            var constructorObj = classGen.createConstructor(text);

            if (!constructorObj.errorFlag) {
                editor.edit(
                    edit => editor.selections.forEach(
                      selection => 
                      {
                        edit.insert(selection.end, constructorObj.generatedCode);
                      }
                    )
                );
            }
        } 
        catch (error) {
            console.log(error);
            vscode.window.showErrorMessage('Something went wrong! Try that the properties are in this format: "Private p_**** As String"');
        }
    });

    let disposableAttributesList = vscode.commands.registerCommand('extension.generateVBAttributesList', function () {
        var editor = vscode.window.activeTextEditor;
        if (!editor)
            return; // No open text editor

        var selection = editor.selection;
        var text = editor.document.getText(selection);

        if (text.length < 1) {
            vscode.window.showErrorMessage('No selected properties.');
            return;
        }

        try {
            var attrListObj = classGen.createAttributesList(text);
            if (!attrListObj.errorFlag) {
                var attrListText = attrListObj.attributesProperty + attrListObj.generatedCode;
    
                editor.edit(
                    edit => editor.selections.forEach(
                      selection => 
                      {
                        edit.insert(selection.end, attrListText);
                      }
                    )
                );
            }
        } 
        catch (error) {
            console.log(error);
            vscode.window.showErrorMessage('Something went wrong! Try that the properties are in this format: "Private p_**** As String"');
        }
    });

    let disposableAttributeFormatList = vscode.commands.registerCommand('extension.generateVBAttributesWithFormatList', function () {
        var editor = vscode.window.activeTextEditor;
        if (!editor)
            return; // No open text editor

        var selection = editor.selection;
        var text = editor.document.getText(selection);

        if (text.length < 1) {
            vscode.window.showErrorMessage('No selected properties.');
            return;
        }

        try {
            var attrListObj = classGen.createAttributesWithFormatList(text);

            if (!attrListObj.errorFlag) {
                var attrListText = attrListObj.attributesProperty + attrListObj.generatedCode;
    
                editor.edit(
                    edit => editor.selections.forEach(
                      selection => 
                      {
                        edit.insert(selection.end, attrListText);
                      }
                    )
                );
            }
        } 
        catch (error) {
            console.log(error);
            vscode.window.showErrorMessage('Something went wrong! Try that the properties are in this format: "Private p_**** As String"');
        }
    });

    let disposableFactory = vscode.commands.registerCommand('extension.generateVBFactory', function () {
        let editor = vscode.window.activeTextEditor;
        if (!editor)
            return; // No open text editor

        const selection = editor.selection;
        let text = editor.document.getText(selection);

        if (text.length < 1) {
            vscode.window.showErrorMessage('No selected properties.');
            return;
        }

        try {
            const filePath = editor.document.fileName;
            let root = path.dirname(filePath);

            const ext = path.extname(filePath);
            let fileName = path.basename(filePath, ext);

            moduleGen.createFactory(text, fileName, root, 'Factories');
        } 
        catch (error) {
            console.log(error);
            vscode.window.showErrorMessage('Something went wrong! Try that the properties are in this format: "Private p_**** As String"');
        }
    });

    let disposableFullProcess = vscode.commands.registerCommand('extension.generateVBFullClass', function() {
        _generateFullClass(false);
    });

    let disposableFullProcessWithFactory = vscode.commands.registerCommand('extension.generateVBFullClassWithFactory', function() {
        _generateFullClass(true);
    });

    context.subscriptions.push(disposableGS);
    context.subscriptions.push(disposableConstructors);
    context.subscriptions.push(disposableAttributesList);
    context.subscriptions.push(disposableAttributeFormatList);
    context.subscriptions.push(disposableFactory);
    context.subscriptions.push(disposableFullProcess);
    context.subscriptions.push(disposableFullProcessWithFactory);
}

// this method is called when your extension is deactivated
function deactivate() { }

function _generateFullClass(factory) {
    let editor = vscode.window.activeTextEditor;
    if (!editor)
    return; // No open text editor

    const selection = editor.selection;
    let text = editor.document.getText(selection);

    if (text.length < 1) {
        vscode.window.showErrorMessage('No selected properties.');
        return;
    }

    try {
        let lines = ``;
        let attrFormatListObj = classGen.createAttributesWithFormatList(text);
        if (attrFormatListObj.errorFlag) return;
        else {
            lines += `${attrFormatListObj.attributesProperty}`;
            lines += `${attrFormatListObj.generatedCode}`;
            text += `${attrFormatListObj.attributesProperty}`;
        }

        let constructorObj = classGen.createConstructor(text);
        if (constructorObj.errorFlag) return;
        else {
            lines += `${constructorObj.generatedCode}`;
        }

        let getterSetterObj = classGen.createGetterAndSetter(text);
        if (getterSetterObj.errorFlag) return;
        else {
            lines += getterSetterObj.generatedCode;
        }

        editor.edit(
            edit => editor.selections.forEach(
                selection => 
                {
                    edit.insert(selection.end, lines);
                }
            )
        );
        
        if (factory) {
            vscode.window.showInformationMessage('Full class created with success! Now generating the factory...');
    
            const filePath = editor.document.fileName;
            let root = path.dirname(filePath);
    
            const ext = path.extname(filePath);
            let fileName = path.basename(filePath, ext);
    
            moduleGen.createFactory(text, fileName, root, 'Factories');
        } else vscode.window.showInformationMessage('Full class created with success!');
    } 
    catch (error) {
        console.log(error);
        vscode.window.showErrorMessage('Something went wrong! Try that the properties are in this format: "Private p_**** As String"');
    }
}

exports.activate = activate;
exports.deactivate = deactivate;