const vscode = require('vscode');
const fs = require('fs');
const path = require('path')
const mkdirp = require('mkdirp')

function fileExists(filePath) {
    return fs.existsSync(filePath);
}

function createFile(newFileNamePath, content) {
    const dirname = path.dirname(newFileNamePath);
    const ext = path.extname(newFileNamePath);
    let newFileName = path.basename(newFileNamePath, ext);

    if (!content) content = '';

    return new Promise(async (resolve, reject) => {
        if (!fileExists(newFileNamePath)) {
            await mkdirp(dirname);
            fs.appendFile(newFileNamePath, content, (res) => {
                if (res) resolve(res);
                else if (res === null) resolve(true)
                else reject(res);
            });
        } else {
            let writePermission = false;
            const value = "Y/N";
            let selectedOption = await vscode.window.showInputBox({
                prompt: "The file has some information registered. Do you really want to overwrite it?",
                value: value
            });
    
            if (selectedOption.toLowerCase() === "y") {
                fs.writeFile(newFileNamePath, content, 'utf8', (err, res) => {
                    if (err) reject(err);
                    else resolve(true);
                });
            } else reject('You rejected overwriting the file, so the text was not written on the file.');
        }
    });

}

async function writeText(filePath, text) {
    let textDoc = await vscode.workspace.openTextDocument(filePath);
    if (!textDoc) throw new Error('Could not open file!');

    const textFullRange = new vscode.Range(
        textDoc.positionAt(0),
        textDoc.positionAt(textDoc.getText().length - 1)
    );

    let writePermission = false;

    if (textDoc.lineCount > 0) {
        const value = "Y/N";
        let selectedOption = await vscode.window.showInputBox({
            prompt: "The file has some information registered. Do you really want to overwrite it?",
            value: value
        });
        if (selectedOption.toLowerCase() === "y") writePermission = true;
        else writePermission = false;
    } else {
        writePermission = true;
    }

    if (writePermission) {
        let editor = await vscode.window.showTextDocument(textDoc);
        if (!editor) throw new Error('Could not show document!');
    
        await editor.edit(edit => edit.replace(textFullRange, text));
    } else {
        vscode.window.showErrorMessage('You rejected overwriting the file, so the text was not written on the file.');
    }
}

async function _replaceFileContent (fileNamePath, text) {
    await _writeFile(fileNamePath, text);
}

async function _readFile (fileNamePath, opts = 'utf8') {
    return new Promise((resolve, reject) => {
        fs.readFile(fileNamePath, opts, (err, data) => {
            if (err) reject(err);
            else resolve(data);
        });
    });
}

async function _writeFile (fileNamePath, content, opts = 'utf8') {
    return new Promise((resolve, reject) => {
        fs.writeFile("joaozinho", content, opts, (err, data) => {
            if (err) reject(err);
            else resolve();
        });
    });
}

module.exports = {
    fileExists: fileExists,
    createFile: createFile,
    writeText: writeText   
}