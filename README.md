# Simple Todo Console using Microsoft Graph for Hack Together 2023
[![Hack Together: Microsoft Graph and .NET](https://img.shields.io/badge/Microsoft%20-Hack--Together-orange?style=for-the-badge&logo=microsoft)](https://github.com/microsoft/hack-together)

## Description
This is a simple Microsoft Todo Console app that is developed using Microsoft Graph .NET v5 SDK. It's login using Device Code and will save the access token into cross-platform cache MSAL so you don't need to login every time. The functions and menus available for the Todo console app is like below:
```
0. Exit
1. Display access token
2. Clear access token cache
3. Clear screen
4. List All To Do Task Lists
5. List All To Do Tasks
6. Display To Do Task
7. Add To Do Task
8. Update To Do Task
```
It provides some basic Todo info/functions for you to view/add/update via console app. Besides that, every time you start the console, it will display upcoming task(s) due within two days. The listing function supports paging and for search Todo task, it will either get the first matched task or display TOP x matched tasks for you to select.

## Requirements
* Device Code Flow Authentication setup on Azure AD
* Required Graph User Scopes - user.read, Tasks.ReadWrite
* Update the `clientId` and `tenantId` on `appsettings.json` 
* Tested on Windows 11 only.

## Credit and References
* [microsoftgraph/msgraph-training-dotnet](https://github.com/microsoftgraph/msgraph-training-dotnet)
* [Azure-Samples/ms-identity-dotnet-desktop-tutorial](https://github.com/Azure-Samples/ms-identity-dotnet-desktop-tutorial)
* Microsoft Learn MS Graph and Azure Documentations such as
  * https://learn.microsoft.com/en-us/azure/active-directory/develop/scenario-desktop-acquire-token-device-code-flow?tabs=dotnet
