# ReadMe

## Description

This is a small test project to test the SetFieldValueByValue method of the `Microsoft.SharePoint.Client.Field` class.

## Prerequirements

- SharePoint Online Site
- SharePoint Online List
- SharePoint Online List Item Field as a managed metadata field
- SharePoint Online List Item
- A user with access to the SharePoint Online Site and List
- An Azure AD App Registration with the following permissions:
  - Sites.ReadWrite.All
  - User.Read.All
  - TermStore.Read.All

## Getting Started

- Copy the config.template.json file to config.json and fill in the values.
- In the TestApplication.cs file, adjust the values for the `listUrl`, `listItemId`, `topicFieldName`, and `newTermGuid` variables.
- Run the project.

## Bug Description

1) When setting a value in a list field with the `SetFieldValueByValue` method of the `Microsoft.SharePoint.Client.Taxonomy` class, the value is empty after the operation even if there was a term value before. No error is thrown and the code runs without any exceptions or warnings.
2) When copying a file from one library to another with the `CopyFile` method of the `Microsoft.SharePoint.Client` class, the file is successfully copied without any error but the term values of the metadata are incorrect. The term values may be empty or even incorrect. Some screenshots from debugging are in the `images` folder. They show the term values before and after the operation. The term values are correct on the original item but incorrect on the copied item.

The code only has a working example for the first bug.
The second has ony a demo function with relevant code that is not yet called in this application. This bug can be reproduced by calling the `CopyFile` method to copy a file with term metadata from one library to another. We copied it from one Site Collection to another Site Collection in the same tenant.

## Bug Analysis

- The error occurs when calling the "SetFieldValueByValue(ListItem, termValue)" function from Microsoft.SharePoint.Client.Taxonomy.
- Additionally, the error occurs when copying a file using CopyFile from Microsoft.SharePoint.Client.
- **The error is specific to the tenant.** This error occurs at Phönix Contact tenant, but not in two of our development tenants.
- It seems highly likely that the error lies with Microsoft SharePoint. Especially when copying a file, I see no way to introduce any potential source of error in our code.
