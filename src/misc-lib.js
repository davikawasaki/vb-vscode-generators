function transformFirstCharToUpperCase(s) {
    return s.charAt(0).toUpperCase() + s.slice(1);
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
            constStatus: false,
            createStatus: false,
            errorLine: ""
        };
        while (p.startsWith(" ")) p = p.substr(1);
        while (p.startsWith("\t")) p = p.substr(1);
        while (p.endsWith(" ")) p = p.slice(0, -1); // Remove all extra whitespaces from end of sentence

        let words = p.split(" ").map(x => x.replace('\r\n', ''));
        
        // if words == 8 (i.e ["Private", "p_*", "As", "Variant", "'", "FORMAT", "NumberFormat", "VALUE", "######.#0"];)
        if (words.length >= 9) {
            propObj.type = words[3];
            propObj.attribute = words[1];
            
            propObj.formatType = words[6]
            propObj.formatValue = words[8]

            propObj.constStatus = false;
            propObj.createStatus = true;
        }
        // if words == 8 (i.e ["p_*", "As", "Variant", "'", "FORMAT", "NumberFormat", "VALUE", "######.#0"];)
        // if words == 6 (i.e ["Const", "p_*", "As", "Integer", "=", "0"];)
        else if (words.length == 8 || words.length == 6) {
            propObj.errorLine = "Declare if the property is Private or Public!";
        }
        // if words == 7 (i.e ["Private", "Const", "p_*", "As", "Integer", "=", "0"];)
        else if (words.length == 7) {
            propObj.type = words[4];
            propObj.attribute = words[2];

            propObj.formatType = ""
            propObj.formatValue = ""

            propObj.constStatus = true;
            propObj.createStatus = true;
        }
        // if words == ["Private", "p_*", "As", "String"];
        else if (words.length == 4) {
            propObj.type = words[3];
            propObj.attribute = words[1];

            propObj.formatType = ""
            propObj.formatValue = ""
            
            propObj.createStatus = true;            
        }
        // if words == ["p_*", "As", "String"];
        else if (words.length == 3) {
            propObj.type = words[2];
            propObj.attribute = words[0];

            propObj.formatType = ""
            propObj.formatValue = ""
            
            propObj.createStatus = true;        
        }
        // if words == ["p_*"];
        else if (words.length) {
            propObj.type = "Variant";
            propObj.attribute = words[0];

            propObj.formatType = ""
            propObj.formatValue = ""

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

module.exports = {
    transformFirstCharToUpperCase: transformFirstCharToUpperCase,
    extractPropertiesArray: extractPropertiesArray
}