const miscLib = require('./misc-lib')
const fileLib = require('./file-lib')
const moment = require('moment')
const vscode = require('vscode')

async function createFactory(propsText, fileName, root, directory) {
    var properties = propsText.split('\r\n').filter(x => x.length > 2);
    var propList = miscLib.extractPropertiesArray(properties);
    var fileNameFactory = `${fileName}Factory`;


    var generatedCode = `
'*****************************************************
'*** @Module ${fileNameFactory}
'*****************************************************
'*** Construct singleton that calls ******************
'*** Init instantiating class fn. ********************
'*****************************************************
'***  Author: <AUTHOR>
'***  Created: ${moment().format("DD/MM/YYYY")}
'***  Last Update: ${moment().format("DD/MM/YYYY")}
'***  Version: 1.0.0
'*****************************************************
`;

    var errorMessage = `Something went wrong (fix them before proceeding) in the following lines: `;
    var constList = propList.filter(p => p.constStatus);
    var errorList = propList.filter(p => !p.createStatus);

    if (errorList.length) {
        let errObjMsg = miscLib.handleErrors(errorList);
        if (errObjMsg.innerErrorsMessage !== "") vscode.window.showErrorMessage(errObjMsg.innerErrorsMessage);
        else vscode.window.showErrorMessage(errorMessage + errObjMsg.innerErrorsLines);
    }
    else if (constList.length === propList.length) {
        vscode.window.showErrorMessage("All properties are constant and don't need to be instantiated! Factory won't be created.");
    }
    else {
        /** @fn Empty Constructor **/
        let codeEmptySignature = 
`
Public Function EmptyConstruct() As ${fileName}`;
        generatedCode += codeEmptySignature
        let codeEmptyBody =
`
\tSet ${fileName}Obj = New ${fileName}
\t${fileName}Obj.Init `;

        for (var i = 0, j = 0; i < propList.length; i++, j++) {
            if (!propList[i].constStatus) {
                let rawAttribute = propList[i].attribute.split('_')[1]
                if (rawAttribute) {
                    codeEmptyBody += `a${miscLib.transformFirstCharToUpperCase(rawAttribute)}:=""`;

                    if (i !== propList.length - 1) {
                        codeEmptyBody += ', ';
                        // Breaking lines (VB editor in excel has character limits)
                        if (j == 5) {
                            j = 0;
                            // Don't break line if the next el is the last and is constant
                            if (i !== propList.length - 2 && !propList[i+1].constStatus) {
                                codeEmptyBody +=
`_ 
`;
                            }
                        }
                    }
                }
                else {
                    console.error('Error in property:',propList[i].attribute)
                    vscode.window.showErrorMessage('Something went wrong! All properties name has to start with a p_! Change ' + propList[i].attribute + ' to p_ATTRIBUTENAME.');
                }
            } else {
                // Remove unnecessary comma and space
                if (i == propList.length - 1) {
                    if (codeEmptyBody.substr(codeEmptyBody.length - 2) == ", ") codeEmptyBody = codeEmptyBody.slice(0, codeEmptyBody.length - 2);
                }
            }
        }

        codeEmptyBody += `
\tSet EmptyConstruct = ${fileName}Obj
End Function`;
        generatedCode += codeEmptyBody;

        /** @fn Constructor **/
        let codeSignature = 
`

Public Function Construct(`;

        for (var i = 0, j = 0; i < propList.length; i++, j++) {
            if (!propList[i].constStatus) {
                let rawAttribute = propList[i].attribute.split('_')[1]
                if (rawAttribute) {
                    codeSignature += `${miscLib.transformFirstCharToUpperCase(rawAttribute)} As ${propList[i].type}`;
                    if (i !== propList.length - 1) {
                        codeSignature += ', ';
                        // Breaking lines (VB editor in excel has character limits)
                        if (j == 5) {
                            j = 0;
                            // Don't break line if the next el is the last and is constant
                            if (i !== propList.length - 2 && !propList[i+1].constStatus) {
                                codeSignature +=
`_ 
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

        codeSignature += `) As ${fileName}`;
        generatedCode += codeSignature
        
        let codeBody =
`
\tSet ${fileName}Obj = New ${fileName}
\t${fileName}Obj.Init `;

        for (var i = 0, j = 0; i < propList.length; i++, j++) {
            if (!propList[i].constStatus) {
                let rawAttribute = propList[i].attribute.split('_')[1]
                if (rawAttribute) {
                    codeBody += `a${miscLib.transformFirstCharToUpperCase(rawAttribute)}:=${rawAttribute}`;

                    if (i !== propList.length - 1) {
                        codeBody += ', ';
                        // Breaking lines (VB editor in excel has character limits)
                        if (j == 5) {
                            j = 0;
                            // Don't break line if the next el is the last and is constant
                            if (i !== propList.length - 2 && !propList[i+1].constStatus) {
                                codeBody +=
`_ 
`;
                            }
                        }
                    }
                }
                else {
                    console.error('Error in property:',propList[i].attribute)
                    vscode.window.showErrorMessage('Something went wrong! All properties name has to start with a p_! Change ' + propList[i].attribute + ' to p_ATTRIBUTENAME.');
                }
            } else {
                // Remove unnecessary comma and space
                if (i == propList.length - 1) {
                    if (codeBody.substr(codeBody.length - 2) == ", ") codeBody = codeBody.slice(0, codeBody.length - 2);
                }
            }
        }

        codeBody += `
\tSet Construct = ${fileName}Obj
End Function`;
        generatedCode += codeBody;

        try {
            const creationStatus = await fileLib.createFile(`${root}\\${directory}\\${fileNameFactory}.bas`, generatedCode);
            if (creationStatus) vscode.window.showInformationMessage(`Factory created/updated with success at ${directory} folder!`);
            else vscode.window.showErrorMessage(`Unknown error at creating ${directory} folder!`);
        } catch (err) {
            if (err) {
                if (typeof err === "string") vscode.window.showErrorMessage(err);
                else vscode.window.showErrorMessage(`Error at creating ${directory} folder: folder access is not permitted or a full path was not provided.`);
            }
            else vscode.window.showErrorMessage(`Unknown error at creating ${directory} folder!`);
        }
    }

}

module.exports = {
    createFactory: createFactory
}