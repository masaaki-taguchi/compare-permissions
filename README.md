# compare-permissions
<p align="center">
  <img src="https://img.shields.io/badge/Salesforce-00a1e0.svg">
  <img src="https://img.shields.io/badge/JavaScript-yellow.svg?logo=JavaScript&logoColor=white">
  <img src="https://img.shields.io/badge/NodeJS-339933.svg?logo=Node.js&logoColor=white">
  <img src="https://img.shields.io/badge/license-MIT-blue.svg">
</p>

[Japanese](./README-ja.md)  

Salesforce's profile and permission set settings exporter to Excel and compare
## Description
Compare permissions is a CLI tool that exports Salesforce's profile and permission set settings to Excel.  
It also supports multiple organizations and makes it possible to compare their settings on Excel.  
The following settings are supported.

* Object Permission
* Field Level Security
* Layout Assignment
* Record Type Visibility
* User Permission
* Application Visibility
* Tab Visibility
* Apex Class Access
* Visualforce Page Access
* Custom Permission
* Custom Metadata Type Access
* Custom Setting Access
* Login IP Range
* Login Hour
* Session Setting
* Password Policy

## Output image
[<img src="./images/sample.png" width="600">](./images/sample.png)

## Installation
Install [Node.js](https://nodejs.org/) and place the following files in any directory.

* compare-permissions.js
* user_config.yaml
* app_config.yaml
* template.xlsx
* package.json
* package-lock.json

Execute the following command to install the required libraries.
```
$ npm ci
```

## Edit config file
Open user_config.yaml and edit the organization name, login URL, user name, and password to match your environment.
```
org:
  - name: (ANY ORGANIZATION NAME 1)
    loginUrl: "https://test.salesforce.com"
    apiVersion : "62.0"
    userName: "(YOUR USER NAME)"
    password: "(YOUR USER PASSWORD)"
#  - name: (ANY ORGANIZATION NAME 2)
#    loginUrl: "https://login.salesforce.com"
#    apiVersion : "62.0"
#    userName: "(YOUR USER NAME)"
#    password: "(YOUR USER PASSWORD)"
```

Edit the profile name and permission set name(label) of the output target. For permission sets, add "ps: true".
```
target:
  - name: "CustomAdmin"
  - name: "CustomStandardUser"
  - name: "SalesUser"
    ps: true
```

Edit the output setting types or order.
```
settingType: [
  "ObjectPermission",
  "FieldLevelSecurity",
  "LayoutAssignment",
  "RecordTypeVisibility",
  "UserPermission",
  "ApplicationVisibility",
  "TabVisibility",
  "ApexClassAccess",
  "ApexPageAccess",
  "CustomPermission",
  "CustomMetadataTypeAccess",
  "CustomSettingAccess",
  "LoginIpRange",
  "LoginHour",
  "SessionSetting",
  "PasswordPolicy"
]
```

Edit the output target objects. If this parameter is not defined, All objects included in the profiles/permission sets are output targets.
```
object: [
  Account, 
  Contact, 
  Opportunity, 
  User, 
]

```

## Usage

Execute a compare-permission with Node.js in a terminal. If there is no option, It uses the default config file. (default is "./user_config.yaml")
```
$ node compare-permissions.js
```
Execute logs are output to screen.
```
[Settings]
  AppConfigPath: app_config.yaml
  TemplateFilePath: template.xlsx
  ResultFilePath: result.xlsx
  TargetProfiles/PermissionSets:
    CustomAdmin
    CustomStandardUser
    SalesUser(PS)
  TargetSettingTypes:
    ObjectPermission
    FieldLevelSecurity
    :
    PasswordPolicy
  TargetObjects:
    Account
    Contact
    Opportunity
    User

[OrgInfo]
  Name:(YOUR ORG NAME)
  LoginUrl:https://login.salesforce.com
  ApiVersion:62.0
  UserName:(YOUR USER NAME)
[Processing profile: CustomAdmin]
  Retrieving base info...
  Retrieving object permissions...
  Retrieving field-level security...
  Retrieving layout assignments...
  Retrieving record-type visibilities...
  Retrieving apex class accesses...
  Retrieving apex page accesses...
  Retrieving user permissions...
  Retrieving application visibilities...
  Retrieving tab visibilities...
  Retrieving login IP ranges...
  Retrieving login hours...
  Retrieving custom permissions...
  Retrieving custom metadata type accesses...
  Retrieving custom setting accesses...
:
[Exporting to an Excel file: result.xlsx]
Done.
```
When complete the execution, the result excel file will be output. (default is "./result.xlsx")

The following options can be used.
```
usage: compare-permissions.js [-options]
    -c <pathname> specifies a config file path (default is ./user_config.yaml)
    -s            don't display logs of the execution
    -h            output usage
````

## Note
- Labels are output as much as possible, but some can not output.
- SessionSetting is output only some of the settings.
- If the connected user's profile is set to High Assurance, the connection will fail.

## License
compare-permissions is licensed under the MIT license.

