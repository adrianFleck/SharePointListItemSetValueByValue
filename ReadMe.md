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