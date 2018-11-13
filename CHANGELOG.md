# Change Log

## v1.0.0
- Feature: Getters and setters generation script
- Feature: Constructor generation script

## v1.1.0
- Bug fixes: End If on constructor, breaking lines
- Feature: Attributes list generation script

## v1.1.1
- Bug fixes: Const properties weren't considerated for getters

## v1.1.2
- Bug fixes: Avoid breaking line in init fn if there aren't any strings left

## v1.1.3
- Version update only for marketplace

## v1.1.4
- Bug fixes: Reduced the args to break a line (10 to 5) and now constructor breaks the line again (v1.1.2 crashed this subfeature)

## v1.1.5
- Feature: Attributes format list generation script
- Bug fixes: Extra whitespace in the end of properties are now removed

## v1.1.6
- Subfeature: Generation with headers

## v1.2.0
- Feature: Singleton factories from class attributes

## v1.2.1
- Bug fix: Adding mkdirp, fs and path npm dependencies to package.json to allow program-flow continuity

## v1.2.2
- Bug fixes: Use of regExp to check sentence structure + error handling in a more concise way
- README updated with more detailed usage description

## v1.3.0
- Bug fixes: Changed breaking lines method in the class generators and regex evaluation cases for constant properties with format
- Features: Full class generator with or without factory

## v1.3.1
- Subfeatures: Empty constructor + author description in factory generated file

## v1.4.0
- Subfeatures: FORMAT colors available (background and foreground)
- Bug fixes: Break code if some error pops up instead of emitting success messages from extension.js main subscriptions