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
            var attrList = createAttributesList(text);

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

    context.subscriptions.push(disposableGS);
    context.subscriptions.push(disposableConstructors);
    context.subscriptions.push(disposableAttributesList);
}

function _transformFirstCharToUpperCase(s) {
    return s.charAt(0).toUpperCase() + s.slice(1);
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
        
        // if words > 7 (i.e ["Private", "Const", "p_*", "As", "Integer", "=", "0"];)
        if (words.length >= 7) {
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

        // Checking for Variant arrays (can't be edited by force, only on initialization - i.e native method 'Private Sub class_initialize()')
        words.forEach(w => {
            if (w.indexOf("()") !== -1) propObj.constStatus = true;
        });

        extProps.push(propObj);
    }

    return extProps;
}

function createGetterAndSetter(propsText) {
    var properties = propsText.split('\r\n').filter(x => x.length > 2);
    var propList = _extractPropertiesArray(properties);

    var generatedCode = 
`
`;
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
            let rawAttribute = p.attribute.split('_')[1];
            if (rawAttribute) {
                // Removal of variant arrays parenthesis
                if (rawAttribute.indexOf("()") !== -1) rawAttribute = rawAttribute.replace("()", "");
                if (p.attribute.indexOf("()") !== -1) p.attribute = p.attribute.replace("()", "");
                let code = 
`
''''''''''''''''''''''
' ${_transformFirstCharToUpperCase(rawAttribute)} property
''''''''''''''''''''''
Public Property Get ${rawAttribute}() As ${p.type}
\t${rawAttribute} = ${p.attribute}
End Property
`;
                generatedCode += code;
    
                if (!p.constStatus) {
                    let code = 
`
Public Property Let ${rawAttribute}(value As ${p.type})
\t${p.attribute} = value
End Property
`;

                    generatedCode += code;
                }
            }
            else {
                vscode.window.showErrorMessage('Something went wrong! All properties name has to start with a p_! Change ' + p.attribute + ' to p_ATTRIBUTENAME.');
            }
        }
    }

    return generatedCode;
}

function createConstructor(propsText) {
    var properties = propsText.split('\r\n').filter(x => x.length > 2);
    var propList = _extractPropertiesArray(properties);

    var generatedCode = 
`
`;

    var errorMessage = `Something went wrong (fix them before proceeding) in the following lines: `;
    var errorList = propList.filter(p => !p.createStatus);

    if (errorList.length) {
        for(let i=0; i < errorList.length; i++) {
            errorMessage += `\n${errorList[i].errorLine}`;
        }
        vscode.window.showErrorMessage(errorMessage);
    }
    else {
        let codeHeader = '';

        for (var i = 0, j = 0; i < propList.length; i++, j++) {
            if (!propList[i].constStatus) {
                let rawAttribute = propList[i].attribute.split('_')[1]
                if (rawAttribute) {
                    codeHeader += `a${_transformFirstCharToUpperCase(rawAttribute)} As ${propList[i].type}`;
                    if (i !== propList.length - 1) {
                        codeHeader += ', ';
                        // Breaking lines (VB editor in excel has character limits)
                        if (j == 10) {
                            j = 0;
                            // Don't break line if the next el is the last and is constant
                            if (i !== propList.length - 2 && propList[i].constStatus) {
                                codeHeader +=
` _ 
`;
                            }
                        }
                    }
                }
                else {
                    vscode.window.showErrorMessage('Something went wrong! All properties name has to start with a p_! Change ' + propList[i].attribute + ' to p_ATTRIBUTENAME.');
                }
            }
            else {
                // Remove unnecessary comma and space
                if (i == propList.length - 1) {
                    if (codeHeader.substr(codeHeader.length - 2) == ", ") codeHeader = codeHeader.slice(0, codeHeader.length - 2);
                }
            }
        }
        
        let codeSignature = 
`
Public Sub Init(${codeHeader})
`;
        let codeBody =
``;
        for (let p of propList) {
            if (!p.constStatus) {
                let rawAttribute = p.attribute.split('_')[1]
                if (rawAttribute) {
                    let codeInnerBody = ``;
    
                    if (p.type == 'Integer' || p.type == 'Double' || p.type == 'Long') {
                        codeInnerBody = 
`\tIf ${p.attribute} = 0 Then
\t\t${p.attribute} = a${_transformFirstCharToUpperCase(rawAttribute)}
\tEnd If
`;
                    }
                    else if (p.type == 'String') {
                        codeInnerBody = 
`\tIf ${p.attribute} = "" Then
\t\t${p.attribute} = a${_transformFirstCharToUpperCase(rawAttribute)}
\tEnd If
`;
                    }
                    // Variant types
                    else {
                        codeInnerBody = 
`\tIf isEmpty(${p.attribute}) Then
\t\t${p.attribute} = a${_transformFirstCharToUpperCase(rawAttribute)}
\tEnd If
`;
                    }
                    
                    codeBody += codeInnerBody;
                }
                else {
                    vscode.window.showErrorMessage('Something went wrong! All properties name has to start with a p_! Change ' + p.attribute + ' to p_ATTRIBUTENAME.');
                }
            }
        }

        let codeFooter = 
`End Sub`;
        generatedCode += codeSignature + codeBody + codeFooter;
            
    }

    return generatedCode;
}

function createAttributesList(propsText) {
    var properties = propsText.split('\r\n').filter(x => x.length > 2);
    var propList = _extractPropertiesArray(properties);

    var generatedCode = 
`
`;

    var errorMessage = `Something went wrong (fix them before proceeding) in the following lines: `;
    var errorList = propList.filter(p => !p.createStatus);

    if (errorList.length) {
        for(let i=0; i < errorList.length; i++) {
            errorMessage += `\n${errorList[i].errorLine}`;
        }
        vscode.window.showErrorMessage(errorMessage);
    }
    else {
        generatedCode += 
`Private p_attributesList() As Variant

Private Sub class_initialize()
\tp_attributesList = Array(`;

        for (let i = 0, j = 0; i < propList.length; i++, j++) {
            let rawAttribute = propList[i].attribute.split('_')[1]
            if (rawAttribute) {
                generatedCode += `"${rawAttribute}"`;
                if (i !== propList.length - 1) {
                    generatedCode += ', ';
                    // Breaking lines (VB editor in excel has character limits)
                    if (j == 10) {
                        j = 0;
                        generatedCode +=
` _ 
`;
                    }
                }
                else {
                    generatedCode += `)
End Sub`;
                }
            }
            else {
                vscode.window.showErrorMessage('Something went wrong! All properties name has to start with a p_! Change ' + propList[i].attribute + ' to p_ATTRIBUTENAME.');
            }
        }
    }

    return generatedCode;
}

exports.activate = activate;

// this method is called when your extension is deactivated
function deactivate() { }

exports.deactivate = deactivate;