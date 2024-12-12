
# SharePoint Online (SPO) API 

The following is interesystems classes to allow the Upload and Download of files to a sharpoint location. This is using the Sharepoint API V1 https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service?tabs=csom



## Authors

- [Mark O'Reilly](https://www.github.com/Sparkei)
- Dean White



## Installation

You can import the code by downloading the Sharepoint Rest API xml SharepointRestAPI.xml into your namespace. Note that this should show all as ADD on your system it is on my systems i already have 2 classes

TODO Insert Picture here of the deploy screen 

TODO Insert Picture of production here .

See OAUTH SETUP PAGES AS NEED TO SET UP OAUTH IN TIE as well

    
## API Reference- Note .. means parameters

#### Get all files in folder- can be filtered  start from /sites/SharepointSitename/

```http
  _api/web/GetFolderByServerRelativeUrl(..folder)/Files/
```

| Parameter | Type     | Description                |
| :-------- | :------- | :------------------------- |
| `$filter` | `string` | Used for datetimemodified i.e. =TimeLastModified gt datetime'2024-12-11T13:43:21.821' |

#### Post file  item

```http
  POST _api/web/GetFolderByServerRelativeUrl(..folder)/Files/add(url='folder'overwrite=..boolean)
```

| Parameter | Type     | Description                       |
| :-------- | :------- | :-------------------------------- |
| `folder`      | `string` | replace with folder name     
| `boolean`      | `string` | replace with true or false   
#### Delete file 
```http
DELETE /api/web/GetFolderByServerRelativeUrl(..SharepointFilePath/Files('..FileName)')

| Parameter | Type     | Description                       |
| :-------- | :------- | :-------------------------------- |
| `FileName`      | `string` | Replace with filename to delete    
```
## Seruices 

SharepointUploadFileEndpoint

A standard EnsLib.File.PassthroughService that passes the file directly to the SharepointRESTConnector Operation and uploads the file to the site.

 SharepointRESTConnector

A custom service for initiating the download process



## Processes
SharepointFileProcessor

Starts off to get a file list from the API, then download these files and uploads to a standard file adapter.

- SharePointService

The service to send all API requests to

- UploadFileToService

the file adapter to upload all files too, after downloaded from SPO

- DeleteFileAfterDownload

Default is do not delete file from Sharepoint site after, if true the file will be deleted from SPO.
##  SharepointDownloadFileEndpoint

A standard EnsLib.File.OutboundAdapter to move the file to.

 SharepointRESTConnector

API calls are here to upload files, getfilelist, download files and delete files.
## OAUTH setup

Server

This needs to setup with the following setting: · Issuer endpoint: https://accounts.accesscontrol.windows.net/{TenantID}/tokens/OAuth/2

· SSL/TLS configuration: Blank TLS1.3

· Authorization endpoint: https://accounts.accesscontrol.windows.net/{TenantID}/tokens/OAuth/2

Client
· Set your Application and Client Name the same, this will be the setting you use in OAuthClientApplicationName of the rest operation

· Put a Tick in USE TLS/SSL and a hostname of localhost

· Select Client Credentials and Form Encoded Body at the bottom of the form

· Select the Client Credentials Tab and set the Client ID as {ClientID}@{TennantId}, then set your Client Secret



TODO Insert OAUTH Page settings here 

 
## Demo

Upload file

Getting file list 

Deleting File


TODO Insert video

