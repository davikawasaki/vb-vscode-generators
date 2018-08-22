const vscode = require('vscode');
const miscLib = require('./misc-lib')

function createGetterAndSetter(propsText) {
    var properties = propsText.split('\r\n').filter(x => x.length > 2);
    var propList = miscLib.extractPropertiesArray(properties);

    var generatedCode = 
`

'***************************
'*** GETTERS AND SETTERS ***
'***************************
`;
    var errorMessage = `Something went wrong (fix them before proceeding) in the following lines: `;
    var errorList = propList.filter(p => !p.createStatus);

    if (errorList.length) {
        let errObjMsg = miscLib.handleErrors(errorList);
        if (errObjMsg.innerErrorsMessage !== "") vscode.window.showErrorMessage(errObjMsg.innerErrorsMessage);
        else vscode.window.showErrorMessage(errorMessage + errObjMsg.innerErrorsLines);
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
' ${miscLib.transformFirstCharToUpperCase(rawAttribute)} property
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
    var propList = miscLib.extractPropertiesArray(properties);

    var generatedCode = 
`

'************************
'*** MAIN CONSTRUCTOR ***
'************************
`;

    var errorMessage = `Something went wrong (fix them before proceeding) in the following lines: `;
    var errorList = propList.filter(p => !p.createStatus);

    if (errorList.length) {
        let errObjMsg = miscLib.handleErrors(errorList);
        if (errObjMsg.innerErrorsMessage !== "") vscode.window.showErrorMessage(errObjMsg.innerErrorsMessage);
        else vscode.window.showErrorMessage(errorMessage + errObjMsg.innerErrorsLines);
    }
    else {
        let codeSignature = 
`
Public Sub Init(`;

        for (var i = 0, j = 0; i < propList.length; i++, j++) {
            if (!propList[i].constStatus) {
                let rawAttribute = propList[i].attribute.split('_')[1]
                if (rawAttribute) {
                    codeSignature += `a${miscLib.transformFirstCharToUpperCase(rawAttribute)} As ${propList[i].type}`;
                    if (i !== propList.length - 1) {
                        codeSignature += ', ';
                        // Breaking lines (VB editor in excel has character limits)
                        if (j == 5) {
                            j = 0;
                            // Don't break line if the next el is the last and is constant
                            if (i !== propList.length - 2 && !propList[i+1].constStatus) {
                                codeSignature +=
` _ 
`;
                            }
                        }
                    }
                }
                else {
                    console.error('Error in property:',propList[i])
                    vscode.window.showErrorMessage('Something went wrong! All properties name has to start with a p_! Change ' + propList[i].attribute + ' to p_ATTRIBUTENAME.');
                }
            }
            else {
                // Remove unnecessary comma and space
                if (i == propList.length - 1) {
                    if (codeSignature.substr(codeSignature.length - 2) == ", ") codeSignature = codeSignature.slice(0, codeSignature.length - 2);
                }
            }
        }

        codeSignature += `)
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
\t\t${p.attribute} = a${miscLib.transformFirstCharToUpperCase(rawAttribute)}
\tEnd If
`;
                    }
                    else if (p.type == 'String') {
                        codeInnerBody = 
`\tIf ${p.attribute} = "" Then
\t\t${p.attribute} = a${miscLib.transformFirstCharToUpperCase(rawAttribute)}
\tEnd If
`;
                    }
                    // Variant types
                    else {
                        codeInnerBody = 
`\tIf isEmpty(${p.attribute}) Then
\t\t${p.attribute} = a${miscLib.transformFirstCharToUpperCase(rawAttribute)}
\tEnd If
`;
                    }
                    
                    codeBody += codeInnerBody;
                }
                else {
                    console.error('Error in property:',p.attribute)
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
    var propList = miscLib.extractPropertiesArray(properties);

    var generatedCode = 
`
`;

    var errorMessage = `Something went wrong (fix them before proceeding) in the following lines: `;
    var errorList = propList.filter(p => !p.createStatus);

    if (errorList.length) {
        let errObjMsg = miscLib.handleErrors(errorList);
        if (errObjMsg.innerErrorsMessage !== "") vscode.window.showErrorMessage(errObjMsg.innerErrorsMessage);
        else vscode.window.showErrorMessage(errorMessage + errObjMsg.innerErrorsLines);
    }
    else {
        generatedCode += 
`Private p_attributesList() As Variant

'***************************************************
'*** CLASS ARRAY LIST INITIALIZATION WITH VALUES ***
'***************************************************

Private Sub class_initialize()
\tp_attributesList = Array(`;

        for (let i = 0, j = 0; i < propList.length; i++, j++) {
            let rawAttribute = propList[i].attribute.split('_')[1]
            if (rawAttribute) {
                generatedCode += `"${rawAttribute}"`;
                if (i !== propList.length - 1) {
                    generatedCode += ', ';
                    // Breaking lines (VB editor in excel has character limits)
                    if (j == 5) {
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

function createAttributesWithFormatList(propsText) {
    var properties = propsText.split('\r\n').filter(x => x.length > 2);
    var propList = miscLib.extractPropertiesArray(properties);

    var generatedCode = 
`
`;

    var errorMessage = `Something went wrong (fix them before proceeding) in the following lines: `;
    var errorList = propList.filter(p => !p.createStatus);

    if (errorList.length) {
        let errObjMsg = miscLib.handleErrors(errorList);
        if (errObjMsg.innerErrorsMessage !== "") vscode.window.showErrorMessage(errObjMsg.innerErrorsMessage);
        else vscode.window.showErrorMessage(errorMessage + errObjMsg.innerErrorsLines);
    }
    else {
        generatedCode += 
`Private p_attributesList() As Variant
Private p_attributesFormatTypesList() As Variant
Private p_attributesFormatValuesList() As Variant

'***************************************************
'*** CLASS ARRAY LIST INITIALIZATION WITH VALUES ***
'***************************************************

Private Sub class_initialize()
\tp_attributesList = Array(`;

        for (let i = 0, j = 0; i < propList.length; i++, j++) {
            let rawAttribute = propList[i].attribute.split('_')[1]
            if (rawAttribute) {
                generatedCode += `"${rawAttribute}"`;
                if (i !== propList.length - 1) {
                    generatedCode += ', ';
                    // Breaking lines (VB editor in excel has character limits)
                    if (j == 5) {
                        j = 0;
                        generatedCode +=
` _ 
`;
                    }
                }
                else {
                    generatedCode += `)`;
                }
            }
        }

        generatedCode += 
`
\tp_attributesFormatTypesList = Array(`;

        for (let i = 0, j = 0; i < propList.length; i++, j++) {
            generatedCode += `"${propList[i].formatType}"`;
            if (i !== propList.length - 1) {
                generatedCode += ', ';
                // Breaking lines (VB editor in excel has character limits)
                if (j == 5) {
                    j = 0;
                    generatedCode +=
` _ 
`;
                }
            }
            else {
                generatedCode += `)`;
            }
        }

        generatedCode += 
`
\tp_attributesFormatValuesList = Array(`;

        for (let i = 0, j = 0; i < propList.length; i++, j++) {
            generatedCode += `"${propList[i].formatValue}"`;
            if (i !== propList.length - 1) {
                generatedCode += ', ';
                // Breaking lines (VB editor in excel has character limits)
                if (j == 5) {
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
    }

    return generatedCode;
}

module.exports = {
    createGetterAndSetter: createGetterAndSetter,
    createConstructor: createConstructor,
    createAttributesList: createAttributesList,
    createAttributesWithFormatList: createAttributesWithFormatList         
}