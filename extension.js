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
            var getterAndSetter = classGen.createGetterAndSetter(text);

            editor.edit(
                edit => editor.selections.forEach(
                  selection => 
                  {
                    edit.insert(selection.end, getterAndSetter);
                  }
                )
              );
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
            var constructor = classGen.createConstructor(text);

            editor.edit(
                edit => editor.selections.forEach(
                  selection => 
                  {
                    edit.insert(selection.end, constructor);
                  }
                )
              );
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
            var attrList = classGen.createAttributesList(text);

            editor.edit(
                edit => editor.selections.forEach(
                  selection => 
                  {
                    edit.insert(selection.end, attrList);
                  }
                )
              );
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
            var attrList = classGen.createAttributesWithFormatList(text);

            editor.edit(
                edit => editor.selections.forEach(
                  selection => 
                  {
                    edit.insert(selection.end, attrList);
                  }
                )
              );
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

    context.subscriptions.push(disposableGS);
    context.subscriptions.push(disposableConstructors);
    context.subscriptions.push(disposableAttributesList);
    context.subscriptions.push(disposableAttributeFormatList);
    context.subscriptions.push(disposableFactory);
}

exports.activate = activate;

// this method is called when your extension is deactivated
function deactivate() { }

exports.deactivate = deactivate;