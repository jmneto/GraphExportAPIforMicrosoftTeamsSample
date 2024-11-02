# Microsoft Graph Export API for Microsoft Teams Messages Sample

# Introduction

This sample application demonstrates how to use the Microsoft Graph Export API to extract and process messages from Microsoft Teams. The application is designed to handle large volumes of data efficiently by utilizing multitasking and batch processing.

# Key Features
1.	**Data Extraction:**  Extract messages from Microsoft Teams using the Microsoft Graph Export API.
2.	**Data Storage:**  Store the extracted messages in an SQL database.
3.	**Logging:**  Log the application's activities to both the console and Azure Storage.
4.	**Configuration:**  Flexible configuration options for input/output storage, logging, and database interactions.
5.	**Multitasking:** Optimize performance by running concurrent tasks for different stages of the data extraction and processing pipeline. 

# Configuration Settings

## Data Input and Output Configuration
- **InputStorageAccountUrl:** URL of the Azure Storage Account containing the input files.  
- **InputStorageConnectionString:** Connection string to the Azure Storage Account containing the input files. Takes precedence over InputStorageAccountUrl.
- **InputStorageContainer:** Container where the input files are stored. Default is inputcontainer. 

## Input Files Wildcards
- **MailBoxesWildCardPattern:** Wildcard pattern for the input files containing the mailboxes to be processed. Default is **AllMailBoxes\*.json**  .  

## Logs Container Configuration
- **LogStorageAccountUrl:** URL of the Azure Storage Account containing the Sample Application logs container location.
- **LogStorageConnectionString:** Connection string to the Azure Storage Account containing the Sample Application logs container location. Takes precedence over LogStorageAccountUrl.
- **LogStorageContainer:** Container where the Sample Application logs are created. Default is logs. 

**Example:** 
```
"InputStorageAccountUrl": "", 
"InputStorageConnectionString": "",
"InputStorageContainer": "inputcontainer",
"MailBoxesWildCardPattern" : "AllMailBoxes*.json",
"LogStorageAccountUrl": "", 
"LogStorageConnectionString": "",
"LogStorageContainer": "logs",
```

## Teams Export API Configuration

- **TenantId:** The Tenant ID of the Azure AD tenant.
- **ApplicationId:** The Application ID of the Azure AD application.
- **ClientSecret:** The Client Secret of the Azure AD application.
- **StartDateTimeUTCString:** The start date for the extraction in UTC format.
- **EndDateTimeUTCString:** The end date for the extraction in UTC format.
- **GetUserMessagesBatchSize:** Batch size for the GetUserMessages API call. Default is 800. 
- **MicrosoftTeamsAPIModel:**  Microsoft Teams API Licensing Model. Possible values are A or B. Default is A. 

**Example:** 
```
"TenantId": "",
"ApplicationId": "",
"ClientSecret": "",
"MicrosoftTeamsAPIModel": "A"
"StartDateTimeUTCString": "2023-01-01Z",
"EndDateTimeUTCString": "2023-12-31Z",
"GetUserMessagesBatchSize": 800,
```

Required permissions for the application to access the Microsoft Teams Export API**

```
Calendars.Read 
Chat.Read.All
User.Read
```

## SQL Database Configuration

- **SqlConnectionString:** The connection string to the SQL Database.
- **ResetDB:** If set to 1, the application will drop and create the tables in the SQL Database, losing all existing data. Default is 0. 

**Example:** 
```
"SqlConnectionString": ""
"ResetDB" : 0
```

SQL Required Permissions:
```
ALTER ROLE [db_Owner] ADD MEMBER [loginaccount] 
ALTER ROLE [db_datareader] ADD MEMBER [loginaccount]  
ALTER ROLE [db_datawriter] ADD MEMBER [loginaccount]  
ALTER ROLE [db_ddladmin]   ADD MEMBER [loginaccount]  
```

## Multitasking

- **MailBoxLoaderTaskLimit:** Maximum number of concurrent tasks for the Mailbox Loader stage. Default is 8.
- **GraphLoaderTaskLimit:** Maximum number of concurrent tasks for the Teams Export API Graph Loader stage. Default is 8.
- **PreProcTaskLimit:** Maximum number of concurrent tasks for the Pre-Processing stage. Default is 16. 

**Example:**
```
"MailBoxLoaderTaskLimit": 8,
"GraphLoaderTaskLimit": 8,
"PreProcTaskLimit": 16,
```

## Logging 

The application logs to the Console Output, which can be seen when running locally or as a WebJob plus a Log File is created in Sample Application Log Container.
The log File name is Logfile_{guid}_index.log. If the application cannot open the log file, it will log to the console an error and abort.

# SQL Tables

## MailBox

Stores the mailboxes imported from AllMailBoxes.json. Controls calls to Graph user APIs.

## UserMessage

Stores the messages and raw JSON resulting from the Teams Export API calls related to users. Only user messages exist in this table.

## GraphCallProcessed

Controls Teams Export API calls progress to avoid redundant calls, allowing the Sample Application to resume if aborted.

# Resetting the Database
 
The ResetDB setting, when enabled, allows the application to recreate necessary tables in the SQL Database, removing all existing data. This setting is only functional in the PRE mode and is essential when database schema updates are required due to new Sample Application versions.

# Conclusion

This sample application provides a comprehensive guide to integrating with the Microsoft Graph Export API for extracting and processing Microsoft Teams messages. By following the provided configurations and understanding the application's structure, developers can efficiently manage and process large volumes of Teams data.

# Sample JSON File

The project includes a sample JSON file to help you understand the structure and format of the data required for the application. This sample file serves as a template for creating your own input files and ensures that the application can correctly parse and process the data.

## Sample JSON File Structure

The sample JSON file, typically named AllMailBoxes.json, contains an array of mailbox objects. Each mailbox object includes essential information such as the mailbox ID, user ID, and other relevant details required for the data extraction process.

**Example:**
```
[
	{
		"DisplayName": "Sachin Tendulkar",
		"ExternalDirectoryObjectId": "b4b488a1-b5e5-4297-8d22-f99ba5cf154e",
		"PrimarySmtpAddress": "Sachin@Sample.com"
	},
	{
		"DisplayName": "Johanna Lorenz",
		"ExternalDirectoryObjectId": "3a3d9d3c-dea8-40c1-9885-30748a2c1635",
		"PrimarySmtpAddress": "JohannaL@Sample.com"
	},
	{
		"DisplayName": "Charlie Brown",
		"ExternalDirectoryObjectId": "48580fcd-1229-4dea-b213-03982d5b3b6e",
		"PrimarySmtpAddress": "Charlie@Sample.com"
	}
]
```