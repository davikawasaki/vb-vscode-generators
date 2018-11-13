function handleErrors(errorList) {
    let innerErrorsMessage = "";
    let innerErrorsLines = "";

    for(let i=0; i < errorList.length; i++) {
        innerErrorsLines += `\n${errorList[i].errorLine}`;
        if (errorList[i].errorMessage && errorList[i].errorMessage !== "") innerErrorsMessage += errorList[i].errorMessage;
    }

    return {
        innerErrorsLines,
        innerErrorsMessage
    }
}

function transformFirstCharToUpperCase(s) {
    return s.charAt(0).toUpperCase() + s.slice(1);
}

function _removeSpaceAndExtraChar(s, spaceBegin, spaceEnd, extraChar, extraCharBegin, extraCharEnd) {
    if (spaceBegin) while (s.startsWith(" ")) s = s.substr(1);
    if (spaceEnd) while (s.endsWith(" ")) s = s.slice(0, -1); // Remove all extra whitespaces from end of sentence
    if (extraChar) {
        if (extraCharBegin) while (s.startsWith(extraChar)) s = s.substr(1);
        if (extraCharEnd) while (s.endsWith(extraChar)) s = s.slice(0, -1);
    }
    
    return s;
}

// i.e ""DSV Charge Code" ' FORMAT NumberFormat VALUE @"
function _checkExtractConstAndFormat(s) {
    s = _removeSpaceAndExtraChar(s, true, true, "\t", true, false);

    let constObj = {
        constStatus: false,
        constValue: null,
        formatObj: null,
        errorCode: null,
        errorMessage: null
    }

    // Search first to check if there is format
    let re = new RegExp("([^\']+)(\\s*\'.+)");
    let searchFormat = re.exec(s);

    if (searchFormat) {
        // Get format details
        if (searchFormat && searchFormat[2].length > 3 && searchFormat[2].indexOf("'") !== -1) {
            let formatObj = _checkExtractFormats(searchFormat[2]);
            if (formatObj.errorCode !== -1) constObj.formatObj = formatObj;
        }
    } else {
        // If there is no format, search only for the value
        re = new RegExp('(.+)');
        let search = re.exec(s);
        
        if (search) {
            constObj.constStatus = true;
            constObj.constValue = search[1];
        } else {
            constObj.errorCode = -1;
            constObj.errorMessage = "A Const value was not found in an attribute! Check it before running the generator.";
        }
    }

    return constObj;
}

// i.e " ' FORMAT NumberFormat VALUE @ FORMATCOLOR BGCOLOR vbRed FGCOLOR vbBlack"
//     " ' FORMAT NumberFormat VALUE @"
//     " ' FORMATCOLOR BGCOLOR vbRed FGCOLOR vbBlack"
function _checkExtractFormats(s) {
    s = _removeSpaceAndExtraChar(s, true, true, "\t", true, false);

    let formatObj = {
        formatType: "",
        formatValue: "",
        formatBgColor: "",
        formatFgColor: "",
        errorCode: null,
        errorMessage: null
    }

    if (s.includes("FORMAT ") && s.includes("FORMATCOLOR ")) {
        let _splitFormatColor = s.split("FORMATCOLOR ");
        if (_splitFormatColor.length !== 2) {
            formatObj.errorCode = -1;
            formatObj.errorMessage = "A FORMAT string constructor has more than two FORMATCOLOR string constructors inside! Check it before running the generator.";
        } else if (!_splitFormatColor[0].includes("FORMAT ")) {
            formatObj.errorCode = -1;
            formatObj.errorMessage = "A FORMAT string constructor has a FORMATCOLOR but does not have a FORMAT constructor! Check it before running the generator.";
        } else {
            formatObj = _extractTypeValueFormat(_splitFormatColor[0], formatObj);
            formatObj = _extractColorFormat(`FORMATCOLOR ${_splitFormatColor[1]}`, formatObj);
        }
    } else if (s.includes("FORMAT ") && !s.includes("FORMATCOLOR ")) {
        if (s.includes("BGCOLOR") || s.includes("FGCOLOR")) {
            formatObj.errorCode = -1;
            formatObj.errorMessage = "A FORMATCOLOR string constructor wasn't constructed properly! Check it before running the generator.";    
        } else {
            formatObj = _extractTypeValueFormat(s, formatObj);
        }
    } else if (s.includes("FORMATCOLOR ") && !s.includes("FORMAT ")) {
        if (s.includes("BGCOLOR") && s.includes("FGCOLOR")) {
            formatObj = _extractColorFormat(s, formatObj);
        } else {
            formatObj.errorCode = -1;
            formatObj.errorMessage = "A FORMATCOLOR string constructor wasn't constructed properly! Check it before running the generator.";    
        }
    } else {
        formatObj.errorCode = -1;
        formatObj.errorMessage = "A FORMATCOLOR string constructor wasn't constructed properly! Check it before running the generator.";
    }

    return formatObj;
}

function _extractTypeValueFormat(s, formatObj) {
    s = _removeSpaceAndExtraChar(s, true, true, "\t", true, false);

    // i.e " ' FORMAT NumberFormat VALUE @ "
    // let re = new RegExp("(\'\\s*)([A-Z]{6})(\\s+)([A-Za-z]+)(\\s+)([A-Z]{5})(\\s+)(.+\\s*)");
    let re = new RegExp("(\\s*)([A-Z]{6})(\\s+)([A-Za-z]+)(\\s+)([A-Z]{5})(\\s+)(.+\\s*)");
    let search = re.exec(s);

    if (search) {
        // Accepts only NumberFormat and NumberFormatLocal
        if (search[2] && search[2] !== "FORMAT") {
            formatObj.errorCode = 2;
            formatObj.errorMessage = "There is a typo in the FORMAT label (" + search[2] + ") ! Check it before running the generator.";
        }
        else if (search[6] && search[6] !== "VALUE") {
            formatObj.errorCode = 6;
            formatObj.errorMessage = "There is a typo in the VALUE label (" + search[6] + ") ! Check it before running the generator.";
        }
        else if (search[4] && search[4].toLowerCase().search("numberformat") === -1) {
            formatObj.errorCode = 4;
            formatObj.errorMessage = "A FORMAT name type needs to be NumberFormat or NumberFormatLocal! Check it before running the generator.";
        }
        else if (search[8] && search[8].search("\"") !== -1) {
            formatObj.errorCode = 8;
            formatObj.errorMessage = "A FORMAT value can not have \"! Remove them before running the generator.";
        }
        else if (search[4] && search[8]) {
            formatObj.formatType = search[4];
            formatObj.formatValue = search[8];
        }
        else {
            formatObj.errorCode = 1;
            formatObj.errorMessage = "A FORMAT type and value were not provided! Check them before running the generator.";
        }
    } else {
        formatObj.errorCode = -1;
        formatObj.errorMessage = "A FORMAT string constructor was not found in an attribute! Check it before running the generator.";

        if (s.indexOf("'") !== -1) {
            let format = s.split("'")[1];
            while (format.startsWith(" ")) format = format.substr(1);

            if (format.length > 3) {
                formatObj.errorCode = 1;
                formatObj.errorMessage = "A FORMAT string was not constructed properly! Check it before running the generator.";
            }
        }
    }

    return formatObj;
}

function _extractColorFormat(s, formatObj) {
    s = _removeSpaceAndExtraChar(s, true, true, "\t", true, false);
    
    // i.e "FORMATCOLOR BGCOLOR 1 FGCOLOR vbBlack"
    let re = new RegExp("(\\s*)([A-Z]{11})(\\s+)([A-Z]{7})(\\s+)([A-Za-z0-9]+)(\\s+)([A-Z]{7})(\\s+)([A-Za-z]+)");
    let search = re.exec(s);

    let reVbConstants = new RegExp("(vb)([A-Z]{1})([a-z]+)");

    if (search) {
        // Accepts only NumberFormat and NumberFormatLocal
        if (search[2] && search[2] !== "FORMATCOLOR") {
            formatObj.errorCode = 2;
            formatObj.errorMessage = "There is a typo in the FORMATCOLOR label (" + search[2] + ") ! Check it before running the generator.";
        }
        else if (search[4] && search[4] !== "BGCOLOR") {
            formatObj.errorCode = 4;
            formatObj.errorMessage = "There is a typo in the BGCOLOR label (" + search[4] + ") ! Check it before running the generator.";
        }
        else if (search[6] && !_isNumber(search[6])) {
            formatObj.errorCode = 6;
            formatObj.errorMessage = "A FORMATCOLOR BGCOLOR value needs to be a number from the palette colors! Check it before running the generator.";
        }
        else if (search[6] && _isNumber(search[6]) && (parseInt(search[6]) < 0 || parseInt(search[6]) > 56)) {
            formatObj.errorCode = 6;
            formatObj.errorMessage = "A FORMATCOLOR BGCOLOR value needs to be a number from the range of palette colors! Check it before running the generator.";
        }
        else if (search[6] && search[6].search("\"") !== -1) {
            formatObj.errorCode = 6;
            formatObj.errorMessage = "A FORMATCOLOR FGCOLOR value can not have \"! Remove them before running the generator.";
        }
        else if (search[8] && search[8] !== "FGCOLOR") {
            formatObj.errorCode = 8;
            formatObj.errorMessage = "There is a typo in the FGCOLOR label (" + search[8] + ") ! Check it before running the generator.";
        }
        else if (search[10] && !reVbConstants.exec(search[10])) {
            formatObj.errorCode = 10;
            formatObj.errorMessage = "A FORMATCOLOR FGCOLOR value needs to be a vb[A-Z]**** type (i.e. vbRed, vbBlack, etc)! Check it before running the generator.";
        }
        else if (search[10] && search[10].search("\"") !== -1) {
            formatObj.errorCode = 10;
            formatObj.errorMessage = "A FORMATCOLOR FGCOLOR value can not have \"! Remove them before running the generator.";
        }
        else if (search[6] && search[10]) {
            formatObj.formatBgColor = search[6];
            formatObj.formatFgColor = search[10];
        }
        else {
            formatObj.errorCode = 1;
            formatObj.errorMessage = "A FORMATCOLOR values were not provided! Check them before running the generator.";
        }
    } else {
        formatObj.errorCode = -1;
        formatObj.errorMessage = "A FORMATCOLOR string constructor was not found in an attribute! Check it before running the generator.";
    }

    return formatObj;
}

function extractPropertiesArray(props) {
    let extProps = [];
    for (let p of props) {
        let propObj = {
            type: "",
            attribute: "",
            rawAttribute: "",
            formatType: "",
            formatValue: "",
            formatBgColor: "",
            formatFgColor: "",
            constStatus: false,
            createStatus: false,
            errorLine: null,
            errorMessage: ""
        };
        
        p = _removeSpaceAndExtraChar(p, true, true, "\t", true, false);

        let words = p.split(" ").map(x => x.replace('\r\n', ''));
        
        if (words.indexOf("Private") === -1 && words.indexOf("Public") === -1) {
            propObj.errorLine = p;
            propObj.errorMessage = "Declare if the property is Private or Public!";
        }
        else if (words.indexOf("Const") !== -1 && words.indexOf("=") !== -1) {
            let constSplit = p.split("=").map(x => x.replace('\r\n', ''));
            if (constSplit.length === 2) {
                let propertyValue = constSplit[0];
                let constValue = constSplit[1];

                propertyValue = _removeSpaceAndExtraChar(propertyValue, false, true, null, false, false);
                constValue = _removeSpaceAndExtraChar(constValue, true, false, null, false, false);

                let constObj = _checkExtractConstAndFormat(constValue);

                propObj.constStatus = constObj.constStatus;

                if (!constObj.errorCode) {
                    if (constObj.formatObj) {
                        if (constObj.formatObj.errorCode) {
                            propObj.errorLine = p;
                            propObj.errorMessage = constObj.formatObj.errorMessage;
                        } else {
                            propObj.formatType = constObj.formatObj.formatType;
                            propObj.formatValue = constObj.formatObj.formatValue;
                            propObj.createStatus = true;
                        }
                    }
                } else {
                    propObj.errorLine = p;
                    propObj.errorMessage = constObj.errorMessage;
                }

                if (!propObj.errorLine) {
                    words = propertyValue.split(" ").map(x => x.replace('\r\n', ''));
    
                    // i.e Private Const
                    if (words.length <= 2) {
                        propObj.errorLine = p;
                        propObj.errorMessage = "No property name was declared in a property! Check it before running the generator.";
                    }
                    // i.e Private Const p_attr
                    else if (words.length === 3) {
                        propObj.type = "Variant";
                        propObj.attribute = words[2];
                        propObj.createStatus = true;
                    }
                    // i.e Private Const p_attr String
                    else if (words.indexOf("As") === -1) {
                        propObj.type = words[3];
                        propObj.attribute = words[2];
    
                        propObj.createStatus = true;
                    }
                    // i.e Private Const p_attr As String
                    else {
                        propObj.type = words[4];
                        propObj.attribute = words[2];
    
                        propObj.createStatus = true;
                    }
                }

            }
            // i.e Private Const p_attr As String
            else {
                propObj.errorLine = p;
                propObj.errorMessage = "A const value was not declared in a property! Check it before running the generator.";
            }
        }
        else if ((words.indexOf("Const") === -1 && words.indexOf("=") !== -1) || (words.indexOf("Const") !== -1 && words.indexOf("=") === -1)) {
            propObj.errorLine = p;
            propObj.errorMessage = "A property has a Const declaration without a value or a value without a Const declaration! Check it before running the generator.";
        }
        else {
            let formatSplit = p.split("'").map(x => x.replace('\r\n', ''));
            
            let propertyValue = formatSplit[0];
            propertyValue = _removeSpaceAndExtraChar(propertyValue, true, true, null, false, false);

            if (formatSplit.length === 2) {
                // let formatValue = "'" + formatSplit[1];
                let formatValue = formatSplit[1];
                formatValue = _removeSpaceAndExtraChar(formatValue, true, true, null, false, false);

                let formatObj = _checkExtractFormats(formatValue);

                if (!formatObj.errorCode) {
                    propObj.formatType = formatObj.formatType;
                    propObj.formatValue = formatObj.formatValue;
                    propObj.formatBgColor = formatObj.formatBgColor;
                    propObj.formatFgColor = formatObj.formatFgColor;
                    propObj.createStatus = true;
                } else {
                    propObj.errorLine = p;
                    propObj.errorMessage = formatObj.errorMessage;
                }

            }
            
            if (!propObj.errorLine) {
                words = propertyValue.split(" ").map(x => x.replace('\r\n', ''));

                // i.e Private 
                if (words.length <= 1) {
                    propObj.errorLine = p;
                    propObj.errorMessage = "No property name was declared in a property! Check it before running the generator.";
                }
                // i.e Private p_attr
                else if (words.length === 2) {
                    propObj.type = "Variant";
                    propObj.attribute = words[1];

                    propObj.createStatus = true;
                }
                // i.e Private p_attr String
                else if (words.indexOf("As") === -1) {
                    propObj.type = words[2];
                    propObj.attribute = words[1];

                    propObj.createStatus = true;
                }
                // i.e. Private p_attr As
                else if (words.indexOf("As") !== -1 && words.length === 3) {
                    propObj.type = "Variant";
                    propObj.attribute = words[1];

                    propObj.createStatus = true;
                }
                // i.e Private p_attr As String
                else {
                    propObj.type = words[3];
                    propObj.attribute = words[1];

                    propObj.createStatus = true;
                }
            }
        }

        // Checking for Variant arrays (can't be edited by force, only on initialization - i.e native method 'Private Sub class_initialize()')
        words.forEach(w => {
            if (w.indexOf("()") !== -1) propObj.constStatus = true;
        });

        extProps.push(propObj);
    }
    
    return extProps;
}

function _isNumber(n) {
    return !isNaN(parseFloat(n)) && isFinite(n);
}

module.exports = {
    handleErrors: handleErrors,
    transformFirstCharToUpperCase: transformFirstCharToUpperCase,
    extractPropertiesArray: extractPropertiesArray
}