# VB Generators
## Extension for Visual Studio Code 
This extension generate VB constructors, getters/setters, class attributes list (with output format types and values) and singleton factories from the VB class variable declarations. You can render them all in a single command too! :)

[![Marketplace Version](https://vsmarketplacebadge.apphb.com/version/davikawasaki.VBGenerators.svg)](https://marketplace.visualstudio.com/items?itemName=davikawasaki.VBGenerators)
[![Installs](https://vsmarketplacebadge.apphb.com/installs/davikawasaki.VBGenerators.svg)](https://marketplace.visualstudio.com/items?itemName=davikawasaki.VBGenerators)
[![Rating](https://vsmarketplacebadge.apphb.com/rating-short/davikawasaki.VBGenerators.svg)](https://marketplace.visualstudio.com/items?itemName=davikawasaki.VBGenerators)

## Table of contents

- [Usage](#usage)
    - [Attributes Structure Recommendations](#attributes-structure-recommendations)
    - [Attribute list with format output](#attribute-list-with-format-output)
- [Examples](#examples)
    - [Rendering Class Constructor](#rendering-class-constructor)
    - [Rendering Attributes Getters and Setters](#rendering-attributes-getters-and-setters)
    - [Rendering Attributes List with class initialization](#rendering-attributes-list-with-class-initialization)
    - [Rendering Attributes List with class initialization and output formats](#rendering-attributes-list-with-class-initialization-and-output-formats)
    - [Rendering Singleton Factory from Class Attributes](#rendering-singleton-factory-from-class-attributes)
    - [Full Rendering Process](#full-rendering-process)
    - [Full Rendering Process With Factory](#full-rendering-process-with-factory)
- [Authors and Collaborators](#authors-and-collaborators)
- [Inspirations](#inspirations)
- [License](#license)
- [Contribution](#contribution)

## Usage

Select the attributes that you want to generate a snippet and run one of the following commands on the command pallete ```Ctrl/Cmd + Shift + P```:

```
$ VB getters and setters
$ VB constructor
$ VB class attribute list
$ VB class attribute list with output format list
$ VB factory from class attributes
$ VB full class
$ VB full class with factory
```

### Attributes Structure Recommendations

1\. Private Const Attribute (needs the **Const** keyword and the attribution in the end):

```vb
Private Const p_attr As String = "ATTRIBUTE"
```

2\. Formatting cell cases (needs the **FORMAT** and **VALUE** keywords, as well as **NumberFormat** or **NumberFormatLocal** as format-property types of a Range object in VBA):

```vb
 ' FORMAT NumberFormat VALUE @ 
 ' FORMAT NumberFormat VALUE yyyy-mm-dd
 ' FORMAT NumberFormat VALUE #####0.#0
```

3\. The number format can be used with a **Const** attribute as well:

```vb
Private Const p_attr As String = "ATTRIBUTE" ' FORMAT NumberFormat VALUE @ 
```

4\. In case you do not say which is the type of the attribute, the extension will understand in all generators that the attribute is a **Variant** type, for instance:

```vb
Private p_attr As
Private p_attr
```

5\. The following cases will output errors, so avoid them at all costs:

```vb
' *** No Public/Private declaration
p_attr As String
' *** No attribution from Const attribute
Private Const p_attr
Private Const p_attr As String = ""
' *** Const attribute with the Const keyword
Private p_attr As String
' *** FORMET instead of FORMAT, TextFormat not acceptable, VALUES instead of VALUE, " usages are not allowed in the format value
Private p_attr As String ' FORMET TextFormat VALUES "@"
```

6\. Factory cases will output the file (if no errors are emitted or there are at least one non-constant attribute) in a Factories/ folder. You may be enquired to approve an override in case the specific factory file already exists in the folder.

### Attribute list with format output

The main idea of having an attribute list with their respective format output is to make it easy to output each attribute value into a Sheet row. The following Sub uses the list and their respective formats to iterate through an object and output all the attributes values into a row:

```vb
'*******************************************
'*** @Sub insertGenericRow *****************
'*******************************************
'*** @Argument {Worksheet} ws **************
'*** @Argument {Variant} classObj **********
'*** @Argument {Integer} myLL **************
'*******************************************
'*** Insert a header/shipment/charge *******
'*** inside a worksheet. *******************
'*******************************************
Sub insertGenericRow(ws As Worksheet, classObj As Variant, ByRef myLL As Integer)
    Dim listLen As Integer
    Dim i As Integer

    i = 1
    listLen = UBound(classObj.attributesList)
    
    ' Iterate through each ordered property from class and send it to the iterated cell with formats
    With ws
        For i = 0 To listLen
            If Not (isEmpty(classObj.attributesFormatTypesList()(i))) Then
                If classObj.attributesFormatTypesList()(i) = "NumberFormat" Then
                    If Not (isEmpty(classObj.attributesFormatValuesList()(i))) Then
                        .Cells(myLL, i + 1).NumberFormat = classObj.attributesFormatValuesList()(i)
                    End If
                End If
            End If

            ' classObj.attributesList() returns the list, and then classObj.attributesList()(i) access an i-element of the list
            .Cells(myLL, i + 1).value = CallByName(classObj, classObj.attributesList()(i), VbGet)
            ' Format cell after inserting into sheet
            
            If Not (isEmpty(classObj.attributesFormatTypesList()(i))) Then
                If classObj.attributesFormatTypesList()(i) = "NumberFormat" Then
                    If Not (isEmpty(classObj.attributesFormatValuesList()(i))) Then
                        .Cells(myLL, i + 1).NumberFormat = classObj.attributesFormatValuesList()(i)
                    End If
                End If
            End If
        Next
    End With

    myLL = myLL + 1
End Sub
```

## Examples
### Rendering Class Constructor
![how use](https://raw.githubusercontent.com/davikawasaki/vb-vscode-generators/master/readme/render_constructor_v1.1.6.gif)

### Rendering Attributes Getters and Setters
![how use](https://raw.githubusercontent.com/davikawasaki/vb-vscode-generators/master/readme/render_getters_setters_v1.1.6.gif)

### Rendering Attributes List with class initialization
![how use](https://raw.githubusercontent.com/davikawasaki/vb-vscode-generators/master/readme/render_attributes_list_v1.1.6.gif)

### Rendering Attributes List with class initialization and output formats
![how use](https://raw.githubusercontent.com/davikawasaki/vb-vscode-generators/master/readme/render_attributes_list_with_formats_v1.1.6.gif)

### Rendering Singleton Factory from Class Attributes
![how use](https://raw.githubusercontent.com/davikawasaki/vb-vscode-generators/master/readme/render_singleton_factory_v1.2.0.gif)

### Full Rendering Process
![how use](https://raw.githubusercontent.com/davikawasaki/vb-vscode-generators/master/readme/render_full_process_v1.3.0.gif)

### Full Rendering Process with Factory
![how use](https://raw.githubusercontent.com/davikawasaki/vb-vscode-generators/master/readme/render_full_process_with_factory_v1.3.0.gif)

## Authors and Collaborators

* Davi Kawasaki

## Inspirations

* [Afmicc - Java Getter and Setter Generator Extension](https://github.com/afmicc/getter-setter-generator)
* [Dkundel - VSCode New File Extension](https://github.com/dkundel/vscode-new-file)

## License
MIT © [davikawasaki](https://github.com/davikawasaki)

## Contribution
Feel free to send me a PR or an issue to improve the code :)
