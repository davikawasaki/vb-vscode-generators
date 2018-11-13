const vscode = require('vscode');
const miscLib = require('./misc-lib')

function createGetterAndSetter(propsText) {
    let errorFlag = false;
    var properties = propsText.split('\r\n').filter(x => {
        return x.length > 2
    });
    var propList = miscLib.extractPropertiesArray(properties);

    var generatedCode = `\r\n\r\n'***************************\r\n'*** GETTERS AND SETTERS ***\r\n'***************************\r\n`;
    var errorMessage = `Something went wrong (fix them before proceeding) in the following lines: `;
    var errorList = propList.filter(p => !p.createStatus);

    if (errorList.length) {
        errorFlag = true;
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
                let code = `\r\n''''''''''''''''''''''\r\n' ${miscLib.transformFirstCharToUpperCase(rawAttribute)} property\r\n''''''''''''''''''''''\r\nPublic Property Get ${rawAttribute}() As ${p.type}\r\n\t${rawAttribute} = ${p.attribute}\r\nEnd Property\r\n`;
                generatedCode += code;
    
                if (!p.constStatus) {
                    let code = `\r\nPublic Property Let ${rawAttribute}(value As ${p.type})\r\n\t${p.attribute} = value\r\nEnd Property\r\n`;
                    generatedCode += code;
                }
            }
            else {
                vscode.window.showErrorMessage('Something went wrong! All properties name has to start with a p_! Change ' + p.attribute + ' to p_ATTRIBUTENAME.');
            }
        }
    }

    return { generatedCode, errorFlag };
}

function createConstructor(propsText) {
    let errorFlag = false;
    var properties = propsText.split('\r\n').filter(x => x.length > 2);
    var propList = miscLib.extractPropertiesArray(properties);

    var generatedCode = `\r\n\r\n'************************\r\n'*** MAIN CONSTRUCTOR ***\r\n'************************\r\n`;

    var errorMessage = `Something went wrong (fix them before proceeding) in the following lines: `;
    var errorList = propList.filter(p => !p.createStatus);

    if (errorList.length) {
        errorFlag = true;
        let errObjMsg = miscLib.handleErrors(errorList);
        if (errObjMsg.innerErrorsMessage !== "") vscode.window.showErrorMessage(errObjMsg.innerErrorsMessage);
        else vscode.window.showErrorMessage(errorMessage + errObjMsg.innerErrorsLines);
    }
    else {
        let codeSignature = `\r\nPublic Sub Init(`;

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
                                codeSignature += ` _ \r\n`;
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

        codeSignature += `)\r\n`;
        
        let codeBody = ``;
        for (let p of propList) {
            if (!p.constStatus) {
                let rawAttribute = p.attribute.split('_')[1]
                if (rawAttribute) {
                    let codeInnerBody = ``;
    
                    if (p.type == 'Integer' || p.type == 'Double' || p.type == 'Long') {
                        codeInnerBody = `\tIf ${p.attribute} = 0 Then\r\n\t\t${p.attribute} = a${miscLib.transformFirstCharToUpperCase(rawAttribute)}\r\n\tEnd If\r\n`;
                    }
                    else if (p.type == 'String') {
                        codeInnerBody = `\tIf ${p.attribute} = "" Then\r\n\t\t${p.attribute} = a${miscLib.transformFirstCharToUpperCase(rawAttribute)}\r\n\tEnd If\r\n`;
                    }
                    // Variant types
                    else {
                        codeInnerBody = `\tIf isEmpty(${p.attribute}) Then\r\n\t\t${p.attribute} = a${miscLib.transformFirstCharToUpperCase(rawAttribute)}\r\n\tEnd If\r\n`;
                    }
                    
                    codeBody += codeInnerBody;
                }
                else {
                    console.error('Error in property:',p.attribute)
                    vscode.window.showErrorMessage('Something went wrong! All properties name has to start with a p_! Change ' + p.attribute + ' to p_ATTRIBUTENAME.');
                }
            }
        }

        let codeFooter = `End Sub`;
        generatedCode += codeSignature + codeBody + codeFooter;
            
    }

    return { generatedCode, errorFlag };
}

function createAttributesList(propsText) {
    var properties = propsText.split('\r\n').filter(x => x.length > 2);
    var propList = miscLib.extractPropertiesArray(properties);

    var attributesProperty = `\r\nPrivate p_attributesList() As Variant\r\n`;

    var generatedCode = `\r\n`;

    var errorMessage = `Something went wrong (fix them before proceeding) in the following lines: `;
    var errorList = propList.filter(p => !p.createStatus);

    var errorFlag = false;

    if (errorList.length) {
        errorFlag = true;
        let errObjMsg = miscLib.handleErrors(errorList);
        if (errObjMsg.innerErrorsMessage !== "") vscode.window.showErrorMessage(errObjMsg.innerErrorsMessage);
        else vscode.window.showErrorMessage(errorMessage + errObjMsg.innerErrorsLines);
    }
    else {
        generatedCode += `\r\n'***************************************************\r\n'*** CLASS ARRAY LIST INITIALIZATION WITH VALUES ***\r\n'***************************************************\r\nPrivate Sub class_initialize()\r\n\tp_attributesList = Array(`;

        for (let i = 0, j = 0; i < propList.length; i++, j++) {
            let rawAttribute = propList[i].attribute.split('_')[1]
            if (rawAttribute) {
                generatedCode += `"${rawAttribute}"`;
                if (i !== propList.length - 1) {
                    generatedCode += ', ';
                    // Breaking lines (VB editor in excel has character limits)
                    if (j == 5) {
                        j = 0;
                        generatedCode += ` _ \r\n`;
                    }
                }
                else {
                    generatedCode += `)\r\nEnd Sub`;
                }
            }
            else {
                vscode.window.showErrorMessage('Something went wrong! All properties name has to start with a p_! Change ' + propList[i].attribute + ' to p_ATTRIBUTENAME.');
            }
        }
    }

    return { generatedCode, attributesProperty, errorFlag };
}

function createAttributesWithFormatList(propsText) {
    var properties = propsText.split('\r\n').filter(x => x.length > 2);
    var propList = miscLib.extractPropertiesArray(properties);

    var attributesProperty = `\r\nPrivate p_attributesList() As Variant\r\nPrivate p_attributesFormatTypesList() As Variant\r\nPrivate p_attributesFormatValuesList() As Variant\r\nPrivate p_attributesFormatBgColorValuesList() As Variant\r\nPrivate p_attributesFormatFgColorValuesList() As Variant\r\n`;

    var generatedCode = `\r\n`;

    var errorMessage = `Something went wrong (fix them before proceeding) in the following lines: `;
    var errorList = propList.filter(p => !p.createStatus);

    var errorFlag = false;

    if (errorList.length) {
        errorFlag = true;
        let errObjMsg = miscLib.handleErrors(errorList);
        if (errObjMsg.innerErrorsMessage !== "") vscode.window.showErrorMessage(errObjMsg.innerErrorsMessage);
        else vscode.window.showErrorMessage(errorMessage + errObjMsg.innerErrorsLines);
    }
    else {
        generatedCode += `\r\n'***************************************************\r\n'*** CLASS ARRAY LIST INITIALIZATION WITH VALUES ***\r\n'***************************************************\r\nPrivate Sub class_initialize()\r\n\tp_attributesList = Array(`;

        for (let i = 0, j = 0; i < propList.length; i++, j++) {
            let rawAttribute = propList[i].attribute.split('_')[1]
            if (rawAttribute) {
                generatedCode += `"${rawAttribute}"`;
                if (i !== propList.length - 1) {
                    generatedCode += ', ';
                    // Breaking lines (VB editor in excel has character limits)
                    if (j == 5) {
                        j = 0;
                        generatedCode += ` _ \r\n`;
                    }
                }
                else {
                    generatedCode += `)`;
                }
            }
        }

        generatedCode += `\r\n\tp_attributesFormatTypesList = Array(`;

        for (let i = 0, j = 0; i < propList.length; i++, j++) {
            generatedCode += `"${propList[i].formatType}"`;
            if (i !== propList.length - 1) {
                generatedCode += ', ';
                // Breaking lines (VB editor in excel has character limits)
                if (j == 5) {
                    j = 0;
                    generatedCode += ` _ \r\n`;
                }
            }
            else {
                generatedCode += `)`;
            }
        }

        generatedCode += `\r\n\tp_attributesFormatValuesList = Array(`;

        for (let i = 0, j = 0; i < propList.length; i++, j++) {
            generatedCode += `"${propList[i].formatValue}"`;
            if (i !== propList.length - 1) {
                generatedCode += ', ';
                // Breaking lines (VB editor in excel has character limits)
                if (j == 5) {
                    j = 0;
                    generatedCode += ` _ \r\n`;
                }
            }
            else {
                generatedCode += `)`;
            }
        }

        generatedCode += `\r\n\tp_attributesFormatBgColorValuesList = Array(`;

        for (let i = 0, j = 0; i < propList.length; i++, j++) {
            // generatedCode += `"${propList[i].formatBgColor}"`;
            if (propList[i].formatBgColor === "") generatedCode += `""`;
            else generatedCode += `${propList[i].formatBgColor}`;
            if (i !== propList.length - 1) {
                generatedCode += ', ';
                // Breaking lines (VB editor in excel has character limits)
                if (j == 5) {
                    j = 0;
                    generatedCode += ` _ \r\n`;
                }
            }
            else {
                generatedCode += `)`;
            }
        }

        generatedCode += `\r\n\tp_attributesFormatFgColorValuesList = Array(`;

        for (let i = 0, j = 0; i < propList.length; i++, j++) {
            generatedCode += `"${propList[i].formatFgColor}"`;
            if (i !== propList.length - 1) {
                generatedCode += ', ';
                // Breaking lines (VB editor in excel has character limits)
                if (j == 5) {
                    j = 0;
                    generatedCode += ` _ \r\n`;
                }
            }
            else {
                generatedCode += `)\r\nEnd Sub`;
            }
        }

    }

    return { generatedCode, attributesProperty, errorFlag };
}

module.exports = {
    createGetterAndSetter: createGetterAndSetter,
    createConstructor: createConstructor,
    createAttributesList: createAttributesList,
    createAttributesWithFormatList: createAttributesWithFormatList         
}