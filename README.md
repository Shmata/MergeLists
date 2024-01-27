# Merge List Columns

This SPFx web part merges SharePoint list columns from different lists in any site into a DetailsList component. According to the administrator's adjustments, either admins or users can configure this web part.

## Contents

- Installation
- Add Web Part
- Wep Part Configuration
- Functionality

## Installation

First, clone the project and open it in VS Code. In the terminal, run the following command:
```bash
  npm run finalBuild
```
![01](https://github.com/Shmata/MergeLists/assets/2398297/71c1466f-02a7-4662-9df5-f768f5c547c6)

Second, navigate to the 'sharepoint\solution' folder to retrieve the 'multi.sppkg' solution package.
![001](https://github.com/Shmata/MergeLists/assets/2398297/287e1194-f392-4856-8550-46a54a7b5b1a)

Then, navigate to a SharePoint site appcatalog and deploy the 'multi.sppkg' solution to your appcatalog site. 
Let’s call this site the destination site.

![3](https://github.com/Shmata/MergeLists/assets/2398297/4759d53e-94f2-4d8e-9bad-0f05409af3e2)


Now, go to the SharePoint Administration Center using this URL: `https://yourtenant-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx#/home`. Since this solution will create a list in the destination site, and the Graph API needs permission to read all sites, it is required to approve an API access request. To do so, access the SharePoint Admin Center, click on `Advanced,` then `API access`. The `Sites.ReadWrite.All` permission is required. 

Finally, select the APIs related to this solution and approve them, as shown in the screenshot below.

![2](https://github.com/Shmata/MergeLists/assets/2398297/74b6de80-3332-47e0-b8b9-10e4f29249f7)

## Add Web Part
After installing the web part, it is ready to use. Navigate to SharePoint `Site contents` and create a page. You can then use this web part on any SharePoint pages. 
Edit the page and add the `Merge` web part.

![4](https://github.com/Shmata/MergeLists/assets/2398297/9a9285f0-ecdf-4e96-bb04-f33c52fcbf97)

## Web Part Configuration 
Once the web part is added to a page, a button will appear in the web part zone. If a site owner or administrator clicks that button, they can easily configure the web part. This button is available only for administrators. All queries will be logged in a SharePoint list, and based on the web part configuration, the administrator can decide who can see the result (Merged columns). Let's explore the different parts of the web part configuration.

![5](https://github.com/Shmata/MergeLists/assets/2398297/d7f051a8-c4f6-4ccf-b820-cc21c50d8137)

As shown in the screenshot above, there are two main parts in the configurations.
   - Styling
     - Description Field
     - Button Alignment
     - Button size

       These three items are related to the style and appearance of the web part
   - Visiblity Status
     - Who can **see** this web part?
     - Who can **edit** this web part?

       Obviously, these two items are related to visibility and edit capability of the web part
       
 
## Functionality 

![6](https://github.com/Shmata/MergeLists/assets/2398297/28d990d0-600e-4181-b6f7-4eacaf3d8f7b)

When an administrator clicks on the main button of the web part, a Fluent UI panel is displayed. Within the panel, there are multi-select Fluent UI dropdown components that administrators can use to select sites, lists, and columns associated with those lists. Upon clicking the **Show grid** button, a Fluent UI DetailsList containing all the selected columns displays accumulated items. The web part provides the ability to filter items.

Once an administrator configures everything and clicks on `Show grid`, all encoded queries will be recorded in the 'MergeLists' list, which is automatically created by the web part. This allows current site users to see the results without any problems.

![7](https://github.com/Shmata/MergeLists/assets/2398297/85bd9253-adc5-4c6f-a66e-767f31464f54)

## Used SharePoint 
Framework Version

![version](https://img.shields.io/badge/version-1.17.4-green.svg)

## Solution

| Solution    | Author                                               |
| ----------- | ------------------------------------------------------- |
| Multi.sppkg | Shahab Matapour |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | January 27, 2024 | Initial release |


## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---
