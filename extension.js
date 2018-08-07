const vscode = require('vscode');

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
            var getterAndSetter = createGetterAndSetter(text);

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
            var constructor = createConstructor(text);

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

    context.subscriptions.push(disposableGS);
    context.subscriptions.push(disposableConstructors);
}

function toPascalCase(str) {
    return str.replace(/\w+/g,w => w[0].toUpperCase() + w.slice(1).toLowerCase());
}

function _extractPropertiesArray(props) {
    let extProps = [];
    for (let p of props) {
        let propObj = {
            type: "",
            attribute: "",
            rawAttribute: "",
            constStatus: false,
            createStatus: false,
            errorLine: ""
        };
        while (p.startsWith(" ")) p = p.substr(1);
        while (p.startsWith("\t")) p = p.substr(1);

        let words = p.split(" ").map(x => x.replace('\r\n', ''));
        
        // if words == ["Private", "Const", "p_*", "As", "Integer", "=", "0"];
        if (words.length == 7) {
            propObj.type = words[4];
            propObj.attribute = words[2];

            propObj.constStatus = true;
            propObj.createStatus = true;
        }
        // if words == ["Private", "p_*", "As", "String"];
        else if (words.length == 4) {
            propObj.type = words[3];
            propObj.attribute = words[1];
            
            propObj.createStatus = true;            
        }
        // if words == ["p_*", "As", "String"];
        else if (words.length == 3) {
            propObj.type = words[2];
            propObj.attribute = words[0];
            
            propObj.createStatus = true;        
        }
        // if words == ["p_*"];
        else if (words.length) {
            propObj.type = "Variant";
            propObj.attribute = words[0];

            propObj.createStatus = true;
        }
        else {
            propObj.errorLine = p;
        }

        extProps.push(propObj);
    }

    return extProps;
}

function createGetterAndSetter(propsText) {
    var properties = propsText.split('\r\n').filter(x => x.length > 2);
    var propList = _extractPropertiesArray(properties);

    var generatedCode = ``;
    var errorMessage = `Something went wrong (fix them before proceeding) in the following lines: `;
    var errorList = propList.filter(p => !p.createStatus);

    if (errorList.length) {
        for(let i=0; i < errorList.length; i++) {
            errorMessage += `\n${errorList[i].errorLine}`;
        }
        vscode.window.showErrorMessage(errorMessage);
    }
    else {
        for (let p of propList) {
            let rawAttribute = p.attribute.split('_')[1]
            if (rawAttribute) {
                let RawAttribute = toPascalCase(rawAttribute);
                let code = 
                    `
                    \t''''''''''''''''''''''
                    \t' get${RawAttribute}() property
                    \t''''''''''''''''''''''
                    \tPublic Property Get get${rawAttribute}() As ${p.type}
                    \t\t${rawAttribute} = ${p.attribute}
                    \tEnd Property
                    `;
                generatedCode += code;
    
                if (!p.constStatus) {
                    let code = 
                        `
                        \tPublic Property Let ${p.attribute}(value As ${p.type})
                        \t\t${p.attribute} = value;
                        \tEnd Property
                        `;
    
                    generatedCode += code;
                }
            }
            else {
                vscode.window.showErrorMessage('Something went wrong! Try that all properties name are in this format: "p_******"');
            }
        }
    }

    return generatedCode;
}

function createConstructor(propsText) {
    var properties = propsText.split('\r\n').filter(x => x.length > 2);
    var propList = _extractPropertiesArray(properties);

    var generatedCode = ``;
    var errorMessage = `Something went wrong (fix them before proceeding) in the following lines: `;
    var errorList = propList.filter(p => !p.createStatus);

    if (errorList.length) {
        for(let i=0; i < errorList.length; i++) {
            errorMessage += `\n${errorList[i].errorLine}`;
        }
        vscode.window.showErrorMessage(errorMessage);
    }
    else {
        let codeSignature = 
            `
            \t'Public Sub Init(
            `;
        let codeBody =
            ``;
        for (let p of propList) {
            if (!p.constStatus) {
                let rawAttribute = p.attribute.split('_')[1]
                if (rawAttribute) {
                    let RawAttribute = toPascalCase(rawAttribute);
                    let codeInnerBody = ``;
    
                    if (p.type == 'Integer' || p.type == 'Double' || p.type == 'Long') {
                        codeInnerBody = 
                        `
                        \tIf ${p.attribute} = 0 Then
                        \t${p.attribute} = a${RawAttribute}
                        `;
                    }
                    else if (p.type == 'String') {
                        codeInnerBody = 
                        `
                        \tIf ${p.attribute} <> "" Then
                        \t${p.attribute} = a${RawAttribute}
                        `;
                    }
                    // Variant types
                    else {
                        codeInnerBody = 
                        `
                        \tIf isEmpty(${p.attribute}) Then
                        \t${p.attribute} = a${RawAttribute}
                        `;
                    }
                    
                    codeBody += codeInnerBody;
                }
                else {
                    vscode.window.showErrorMessage('Something went wrong! Try that all properties name are in this format: "p_******"');
                }
            }
        }

        generatedCode += codeSignature + codeBody + `\tEnd Sub`;
            
    }

    return generatedCode;
}

exports.activate = activate;

// this method is called when your extension is deactivated
function deactivate() { }

exports.deactivate = deactivate;