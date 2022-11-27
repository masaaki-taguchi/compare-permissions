'use strict';

const jsforce = require('jsforce');
const xlsx = require('xlsx-populate');

const METADATA_TRUE = 'true';
const METADATA_FALSE = 'false';
const LABEL_KEY_NAME = '0_Label';
const PREFIX_PROFILE_NAME = '1_';
const PREFIX_PERMISSION_SET_NAME = '2_';
const ADD_INFO_HEADER = 'ADD_INFO:';
const BOOLEAN_OUTPUT = '1';
const BOOLEAN_NO_OUTPUT = '2';
const BOOLEAN_OUTPUT_WITH_NA = '3';
const STYLE_FILL = 'fill';

const OBJECT_PERMISSION = 'ObjectPermission';
const FIELD_LEVEL_SECURITY = 'FieldLevelSecurity';
const LAYOUT_ASSIGNMENT = 'LayoutAssignment';
const RECORD_TYPE_VISIBILITY = 'RecordTypeVisibility';
const USER_PERMISSION = 'UserPermission';
const APPLICATION_VISIBILITY = 'ApplicationVisibility';
const TAB_VISIBILITY = 'TabVisibility';
const APEX_CLASS_ACCESS = 'ApexClassAccess';
const APEX_PAGE_ACCESS = 'ApexPageAccess';
const CUSTOM_PERMISSION = 'CustomPermission';
const LOGIN_IP_RANGE = 'LoginIpRange';
const SESSION_SETTING = 'SessionSetting';
const PASSWORD_POLICY = 'PasswordPolicy';

const METADATA_TYPE_PROFILE = 'Profile';
const METADATA_TYPE_PERMISSION_SET = 'PermissionSet';
const METADATA_TYPE_SESSION_SETTING = 'ProfileSessionSetting';
const METADATA_TYPE_PASSWORD_POLICY = 'ProfilePasswordPolicy';
const KEY_BASE_INFO_NAME = 'name';
const KEY_BASE_INFO_PERMISSION_SET = 'permissionSet';
const KEY_BASE_INFO_CUSTOM = 'custom';
const KEY_BASE_INFO_USER_LICENSE = 'userLicense';
const KEY_BASE_INFO_DESCRIPTION = 'description';
const KEY_SESSION_SETTING_FORCE_LOGOUT = 'forceLogout';
const KEY_SESSION_SETTING_REQUIRED_SESSION_LEVEL = 'requiredSessionLevel';
const KEY_SESSION_SETTING_SESSION_PERSISTENCE = 'sessionPersistence';
const KEY_SESSION_SETTING_SESSION_TIMEOUT = 'sessionTimeout';
const KEY_SESSION_SETTING_SESSION_TIMEOUT_WARNING = 'sessionTimeoutWarning';
const KEY_PASSWORD_POLICY_FORGOT_PASSWORD_REDIRECT = 'forgotPasswordRedirect';
const KEY_PASSWORD_POLICY_LOCKOUT_INTERVAL = 'lockoutInterval';
const KEY_PASSWORD_POLICY_MAX_LOGIN_ATTEMPTS = 'maxLoginAttempts';
const KEY_PASSWORD_POLICY_MINIMUM_PASSWORD_LENGTH = 'minimumPasswordLength';
const KEY_PASSWORD_POLICY_MINIMUM_PASSWORD_LIFE_TIME = 'minimumPasswordLifetime';
const KEY_PASSWORD_POLICY_OBSCURE = 'obscure';
const KEY_PASSWORD_POLICY_PASSWORD_COMPLEXITY = 'passwordComplexity';
const KEY_PASSWORD_POLICY_PASSWORD_EXPIRATION = 'passwordExpiration';
const KEY_PASSWORD_POLICY_PASSWORD_HISTORY = 'passwordHistory';
const KEY_PASSWORD_POLICY_PASSWORD_QUESTION = 'passwordQuestion';

const DEFAULT_USER_CONFIG_PATH = './user_config.yaml';
const DEFAULT_APP_CONFIG_PATH = './app_config.yaml';
const COMMAND_OPTION_SILENT = '-s';
const COMMAND_OPTION_HELP = '-h';
const COMMAND_OPTION_CONFIG = '-c';

let userConfigPath = DEFAULT_USER_CONFIG_PATH;
global.enabledLogging = true;

// analyzes command line options
for (let i = 2; i < process.argv.length; i++) {
  if (process.argv[i] === COMMAND_OPTION_SILENT) {
    global.enabledLogging = false;
  }
  if (process.argv[i] === COMMAND_OPTION_CONFIG) {
    if (i + 1 >= process.argv.length) {
      usage();
    }
    userConfigPath = process.argv[i + 1];
  }
  if (process.argv[i] === COMMAND_OPTION_HELP) {
    usage();
  }
}

loadUserConfig(userConfigPath);
loadAppConfig();

const userConfig = global.userConfig;
const appConfig = global.appConfig;
logging('Settings:');
logging('  AppConfigPath:' + userConfig.appConfigPath);
logging('  TemplateFilePath:' + userConfig.templateFilePath);
logging('  ResultFilePath:' + userConfig.resultFilePath);
logging('  ExcelFormatCopy:' + userConfig.excelFormatCopy);
let target = '';
userConfig.target.forEach(function(value) {
  target += value.name;
  if (value.ps) {target += '(PS)'}
  target += ',';
});
target = target.replace(/,\s*$/, "");
logging('  TargetProfiles/PermissionSets:' + target);
logging('  TargetSettingTypes:' + userConfig.settingType);
logging('  TargetObjects:' + userConfig.object);

(async () => {
  const baseInfoMap = new Map();
  const applicationVisibilityMap = new Map();
  const apexClassAccessMap = new Map();
  const apexPageAccessMap = new Map();
  const objectPermissionMap = new Map();
  const fieldLevelSecurityMap = new Map();
  const fieldLevelSecurityFieldSet = new Set();
  const tabVisibilityMap = new Map();
  const recordTypeVisibilityMap = new Map();
  const loginIpRangeMap = new Map();
  const userPermissionMap = new Map();
  const customPermissionMap = new Map();
  const layoutAssignmentMap = new Map();
  const sessionSettingMap = new Map();
  const passwordPolicyMap = new Map();

  for (const org of global.orgList) {
    const orgName = org.name;
    logging('**** Start to retrieve ****');
    logging('OrgInfo:');
    logging('  Name:' + org.name)
    logging('  LoginUrl:' + org.loginUrl)
    logging('  ApiVersion:' + org.apiVersion)
    logging('  UserName:' + org.userName);
    let conn = new jsforce.Connection({ loginUrl: org.loginUrl, version: org.apiVersion });
    // login to salesforce
    await conn.login(org.userName, org.password, function (err, userInfo) {
      if (err) {
        console.error(err);
        process.exit(1);
      }
    });

    // retrieve metadata in profiles
    if (global.profileNameList.length !== 0) {
      for await (const value of global.profileNameList) {
        const profileNameWorkList = [];
        profileNameWorkList.push(value);
        await conn.metadata.read(METADATA_TYPE_PROFILE, profileNameWorkList, function (err, metadataList) {
          if (err) {
            console.error(err);
            process.exit(1);
          }
          if (!Array.isArray(metadataList)) {
            metadataList = [metadataList];
          }
          if (metadataList[0].fullName === undefined) {
            logging('[Warning]Not exist a profile. Profile : ' + value);
            return;
          }

          retrieveBaseInfo(metadataList, orgName, true, baseInfoMap);
          retrieveObjectPermissions(metadataList, orgName, true, objectPermissionMap);
          retrievefieldLevelSecurities(metadataList, orgName, true, fieldLevelSecurityMap, fieldLevelSecurityFieldSet);
          retrieveLayoutAssignments(metadataList, orgName, true, layoutAssignmentMap);
          retrieveRecordTypeVisibilities(metadataList, orgName, true, recordTypeVisibilityMap);
          retrieveApexClassAccesses(metadataList, orgName, true, apexClassAccessMap);
          retrieveApexPageAccesses(metadataList, orgName, true, apexPageAccessMap);
          retrieveUserPermissions(metadataList, orgName, true, userPermissionMap);
          retrieveApplicationVisibilities(metadataList, orgName, true, applicationVisibilityMap);
          retrieveTabVisibilities(metadataList, orgName, true, tabVisibilityMap);
          retrieveLoginIpRanges(metadataList, orgName, true, loginIpRangeMap);
          retrieveCustomPermissions(metadataList, orgName, true, customPermissionMap);
        });
      }
    }

    const profileNameMap = new Map();
    for (const value of global.profileNameList) {
      const lowerName = value.toLowerCase(value);
      profileNameMap.set(lowerName, value);
    }

    // retrieve metadata in profile session setting
    if (global.profileNameList.length !== 0 && isExecutable(SESSION_SETTING)) {
      const types = [{ type: METADATA_TYPE_SESSION_SETTING, folder: null }];
      const sessionSettingList = [];
      await conn.metadata.list(types, org.apiVersion, function (err, metadataList) {
        if (err) {
          console.error(err);
          process.exit(1);
        }
        if (!Array.isArray(metadataList)) {
          metadataList = [metadataList];
        }
        for (const metadata of metadataList) {
          sessionSettingList.push(metadata.fullName);
        }
      });

      for await (const value of sessionSettingList) {
        const sessionSettingWorkList = [];
        sessionSettingWorkList.push(value);
        await conn.metadata.read(METADATA_TYPE_SESSION_SETTING, sessionSettingWorkList, function (err, metadataList) {
          if (err) {
            console.error(err);
            process.exit(1);
          }
          if (!Array.isArray(metadataList)) {
            metadataList = [metadataList];
          }
          if (metadataList[0].fullName === undefined) {
            logging('[Warning]Not exist a profile. Profile : ' + value);
            return;
          }
          const profileName = metadataList[0].profile;
          if (profileNameMap.has(profileName)) {
            retrieveSessionSetting(metadataList, orgName, profileNameMap.get(profileName), true, sessionSettingMap);
          }
        });
      }
    }

    // retrieve metadata in profile password policy
    if (global.profileNameList.length !== 0 && isExecutable(PASSWORD_POLICY)) {
      const types = [{ type: METADATA_TYPE_PASSWORD_POLICY, folder: null }];
      const passwordPolicyList = [];
      await conn.metadata.list(types, org.apiVersion, function (err, metadataList) {
        if (err) {
          console.error(err);
          process.exit(1);
        }
        if (!Array.isArray(metadataList)) {
          metadataList = [metadataList];
        }
        for (const metadata of metadataList) {
          passwordPolicyList.push(metadata.fullName);
        }
      });

      for await (const value of passwordPolicyList) {
        const passwordPolicyWorkList = [];
        passwordPolicyWorkList.push(value);
        await conn.metadata.read(METADATA_TYPE_PASSWORD_POLICY, passwordPolicyWorkList, function (err, metadataList) {
          if (err) {
            console.error(err);
            process.exit(1);
          }
          if (!Array.isArray(metadataList)) {
            metadataList = [metadataList];
          }
          if (metadataList[0].fullName === undefined) {
            logging('[Warning]Not exist a profile. Profile : ' + value);
            return;
          }
          const profileName = metadataList[0].profile;
          if (profileNameMap.has(profileName)) {
            retrievePasswordPolicy(metadataList, orgName, profileNameMap.get(profileName), true, passwordPolicyMap);
          }
        });
      }
    }

    if (global.permissionSetNameList.length !== 0) {
      // retrieve metadata in permission set
      const permissionSetAPINameList = await getPermissionSetAPINames(conn);
      for await (const value of permissionSetAPINameList) {
        const permissionSetAPINameWorkList = [];
        permissionSetAPINameWorkList.push(value);

        await conn.metadata.read(METADATA_TYPE_PERMISSION_SET, permissionSetAPINameWorkList, function (err, metadataList) {
          if (err) {
            console.error(err);
            process.exit(1);
          }
          if (!Array.isArray(metadataList)) {
            metadataList = [metadataList];
          }

          retrieveBaseInfo(metadataList, orgName, false, baseInfoMap);
          retrieveObjectPermissions(metadataList, orgName, false, objectPermissionMap);
          retrievefieldLevelSecurities(metadataList, orgName, false, fieldLevelSecurityMap, fieldLevelSecurityFieldSet);
          retrieveLayoutAssignments(metadataList, orgName, false, layoutAssignmentMap);
          retrieveRecordTypeVisibilities(metadataList, orgName, false, recordTypeVisibilityMap);
          retrieveApexClassAccesses(metadataList, orgName, false, apexClassAccessMap);
          retrieveApexPageAccesses(metadataList, orgName, false, apexPageAccessMap);
          retrieveUserPermissions(metadataList, orgName, false, userPermissionMap);
          retrieveApplicationVisibilities(metadataList, orgName, false, applicationVisibilityMap);
          retrieveTabVisibilities(metadataList, orgName, false, tabVisibilityMap);
          retrieveCustomPermissions(metadataList, orgName, false, customPermissionMap);
        });
      }
    }

    await compensateObjectsAndFields(conn, objectPermissionMap, fieldLevelSecurityMap, fieldLevelSecurityFieldSet);
    await compensateLayoutAssignments(conn, layoutAssignmentMap);
    await compensateRecordTypeVisibilities(conn, recordTypeVisibilityMap);
    await compensateApexClassAccesses(conn, apexClassAccessMap);
    await compensateApexPageAccesses(conn, apexPageAccessMap);
    await compensateUserPermissions(conn, userPermissionMap);
    await compensateApplicationVisibilities(conn, applicationVisibilityMap);
    await compensateTabVisibilities(conn, tabVisibilityMap);
    await compensateCustomPermissions(conn, customPermissionMap);
    await compensateSessionSetting(conn, sessionSettingMap);
    await compensatePasswordPolicy(conn, passwordPolicyMap);

    await conn.logout(function(err) {
      if (err) { return console.error(err); }
      logging('Logout. userName:' + org.userName);
    });

  }

  // export to an excel file
  let workBook = await xlsx.fromFileAsync(userConfig.templateFilePath);
  logging('Export to an excel file.');
  const sheet = workBook.sheet(0);
  let cell;
  const templateStyleList = getTemplateStyleList(sheet);

  let targetNameWorkX = appConfig.targetNamePosition[0];
  const targetNameWorkY = appConfig.targetNamePosition[1];
  let targetPermissionSetWorkX = appConfig.targetPermissionSetPosition[0];
  const targetPermissionSetWorkY = appConfig.targetPermissionSetPosition[1];
  let targetLicenseWorkX = appConfig.targetLicensePosition[0];
  const targetLicenseWorkY = appConfig.targetLicensePosition[1];
   let targetCustomWorkX = appConfig.targetCustomPosition[0];
  const targetCustomWorkY = appConfig.targetCustomPosition[1];

  // output base information for profile and permission set
  for (let targetCnt = 0; targetCnt < global.targetNameList.length; targetCnt++) {
    let targetName = global.targetNameList[targetCnt];
    let color = appConfig.targetGroupColorDefault;
    if (global.orgList.length >= 2) {
      if (targetCnt % 2 === 0) {
        color = appConfig.targetGroupColor1;
      } else {
        color = appConfig.targetGroupColor2;
      }
    }
    for (const org of global.orgList) {
      const key = org.name + '.' + targetName;
      const value = baseInfoMap.get(key);
      if (value) {
        cell = sheet.row(targetNameWorkY).cell(targetNameWorkX);
        cell.value(value.get(KEY_BASE_INFO_NAME) + '\n(' + org.name + ')');
        cell.style(STYLE_FILL, color);

        cell = sheet.row(targetPermissionSetWorkY).cell(targetPermissionSetWorkX);
        cell.value(convertBoolean(value.get(KEY_BASE_INFO_PERMISSION_SET)));
        cell.style(STYLE_FILL, color);

        cell = sheet.row(targetLicenseWorkY).cell(targetLicenseWorkX);
        cell.value(value.get(KEY_BASE_INFO_USER_LICENSE));
        cell.style(STYLE_FILL, color);

        cell = sheet.row(targetCustomWorkY).cell(targetCustomWorkX);
        cell.value(convertBoolean(value.get(KEY_BASE_INFO_CUSTOM)));
        if (isTrue(value.get(KEY_BASE_INFO_PERMISSION_SET))) {
          cell.style(STYLE_FILL, appConfig.notApplicableColor);
        } else {
          cell.style(STYLE_FILL, color);
        }
      } else {
        cell = sheet.row(targetNameWorkY).cell(targetNameWorkX);
        cell.value(targetName.slice(2) + '\n(' + org.name + ')');
        cell.style(STYLE_FILL, appConfig.notApplicableColor);

        cell = sheet.row(targetPermissionSetWorkY).cell(targetPermissionSetWorkX);
        if (targetName.slice(0, 2) === PREFIX_PERMISSION_SET_NAME) {
          cell.value(convertBoolean(METADATA_TRUE));
        } else {
          cell.value(convertBoolean(METADATA_FALSE));
        }
        cell.style(STYLE_FILL, appConfig.notApplicableColor);

        cell = sheet.row(targetLicenseWorkY).cell(targetLicenseWorkX);
        cell.value('-');
        cell.style(STYLE_FILL, appConfig.notApplicableColor);

        cell = sheet.row(targetCustomWorkY).cell(targetCustomWorkX);
        cell.value('-');
        cell.style(STYLE_FILL, appConfig.notApplicableColor);
      }
      targetNameWorkX++;
      targetPermissionSetWorkX++;
      targetLicenseWorkX++;
      targetCustomWorkX++;
    }

    const typeWorkX = appConfig.typePosition[0];
    let typeWorkY = appConfig.typePosition[1];
    const nameWorkX = appConfig.namePosition[0];
    let nameWorkY = appConfig.namePosition[1];
    const secondNameWorkX = appConfig.secondNamePosition[0];
    let secondNameWorkY = appConfig.secondNamePosition[1];
    const labelWorkX = appConfig.labelPosition[0];
    let labelWorkY = appConfig.labelPosition[1];
    let resultWorkY = appConfig.resultPosition[1];
    let resultWorkX = 0;

    // define the output method for each metadata
    const defaultPutMap = new Map([
      [OBJECT_PERMISSION,
        [OBJECT_PERMISSION, objectPermissionMap, BOOLEAN_NO_OUTPUT,
          'CRUD', 'R', '', true, false]],
      [FIELD_LEVEL_SECURITY,
        [FIELD_LEVEL_SECURITY, fieldLevelSecurityMap, BOOLEAN_NO_OUTPUT,
          'RU', 'R', '', true, false]],
      [LAYOUT_ASSIGNMENT,
        [LAYOUT_ASSIGNMENT, layoutAssignmentMap, BOOLEAN_OUTPUT,
          METADATA_TRUE, '', METADATA_FALSE, false, true]],
      [RECORD_TYPE_VISIBILITY,
        [RECORD_TYPE_VISIBILITY, recordTypeVisibilityMap, BOOLEAN_NO_OUTPUT,
          appConfig.recordTypeVisibilityLabel.visible, '', '',
          true, false]],
      [APEX_CLASS_ACCESS,
        [APEX_CLASS_ACCESS, apexClassAccessMap, BOOLEAN_OUTPUT_WITH_NA,
          METADATA_TRUE, '', METADATA_FALSE, false, false]],
      [APEX_PAGE_ACCESS,
        [APEX_PAGE_ACCESS, apexPageAccessMap, BOOLEAN_OUTPUT_WITH_NA,
          METADATA_TRUE, '', METADATA_FALSE, false, false]],
      [USER_PERMISSION,
        [USER_PERMISSION, userPermissionMap, BOOLEAN_OUTPUT,
          METADATA_TRUE, '', METADATA_FALSE, false, false]],
      [APPLICATION_VISIBILITY,
        [APPLICATION_VISIBILITY, applicationVisibilityMap, BOOLEAN_NO_OUTPUT,
          appConfig.applicationVisibilityLabel.visible, '', '', true, false]],
      [TAB_VISIBILITY,
        [TAB_VISIBILITY, tabVisibilityMap, BOOLEAN_NO_OUTPUT,
          '^' + appConfig.tabVisibilityLabel.defaultOn + '|' + appConfig.tabVisibilityLabel.available,
          '^' + appConfig.tabVisibilityLabel.defaultOff + '|' + appConfig.tabVisibilityLabel.visible,
          appConfig.tabVisibilityLabel.hidden,
          true, false]],
      [LOGIN_IP_RANGE,
        [LOGIN_IP_RANGE, loginIpRangeMap, BOOLEAN_OUTPUT,
          METADATA_TRUE, '', METADATA_FALSE, false, true]],
      [CUSTOM_PERMISSION,
        [CUSTOM_PERMISSION, customPermissionMap, BOOLEAN_OUTPUT,
          METADATA_TRUE, '', METADATA_FALSE, false, false]],
      [SESSION_SETTING,
        [SESSION_SETTING, sessionSettingMap, BOOLEAN_NO_OUTPUT,
          '', '', '', false, true]],
      [PASSWORD_POLICY,
        [PASSWORD_POLICY, passwordPolicyMap, BOOLEAN_NO_OUTPUT,
          '', '', '', false, true]]
    ]);

    // get the type of metadata to output
    const putList = [];
    if (global.settingTypeSet.size === 0) {
      for (const value of defaultPutMap.keys()) {
        putList.push(defaultPutMap.get(value));
      }
    } else {
      for (const value of global.settingTypeSet) {
        if (!defaultPutMap.has(value)) {
          logging('[Error]Incorrect a metadata type name. Name : ' + value);
          process.exit(1);
        }
        putList.push(defaultPutMap.get(value));
      }
    }

    const settingTypeLabelMap = global.appConfig.settingTypeLabel;

    putList.forEach(function (value, index) {
      const settingType = value[0];
      const metadataMap = value[1];
      const sortedMap = new Map([...metadataMap.entries()].sort());
      const booleanBehavior = value[2];
      const fullAuthorityValue = value[3];
      const partialAuthorityValue = value[4];
      const noAuthorityValue = value[5];
      const fillsBlankWithNoAuthorityColor = value[6];
      const notApplicablePermmisonSet = value[7];

      let paintedFirstframe = false;

      for (const key of sortedMap.keys()) {
        if (userConfig.excelFormatCopy === undefined || userConfig.excelFormatCopy === true) {
          putTemplateStyle(sheet, templateStyleList, resultWorkY);
        }
        if (!paintedFirstframe) {
          putFirstFrameStyle(sheet, templateStyleList, resultWorkY);
          paintedFirstframe = true;
        }
        cell = sheet.row(typeWorkY).cell(typeWorkX);
        const settingTypeLabel = settingTypeLabelMap[settingType];
        if (settingTypeLabel !== undefined) {
          cell.value(settingTypeLabel);
        } else {
          cell.value(settingType);
        }
        cell = sheet.row(nameWorkY).cell(nameWorkX);
        const nameRegex = new RegExp('^(.+?)( ' + ADD_INFO_HEADER + '(.+?))*$').exec(key);
        cell.value(nameRegex[1]);
        cell = sheet.row(secondNameWorkY).cell(secondNameWorkX);
        if (nameRegex[3]) {
          cell.value(nameRegex[3]);
        }

        const valueMap = metadataMap.get(key);
        if (valueMap.has(LABEL_KEY_NAME)) {
          const value = valueMap.get(LABEL_KEY_NAME);
          cell = sheet.row(labelWorkY).cell(labelWorkX);
          cell.value(value);
        }
        resultWorkX = appConfig.resultPosition[0];
        for (const targetName of global.targetNameList) {
          let beforeValue;
          let value;
          for (let orgCnt = 0; orgCnt < global.orgList.length; orgCnt++) {
            const org = global.orgList[orgCnt];
            cell = sheet.row(resultWorkY).cell(resultWorkX);
            const key = org.name + '.' + targetName;
            if (valueMap.has(key)) {
              value = valueMap.get(key);
              if (value !== undefined) {
                if (fullAuthorityValue.length > 0 && value.match(new RegExp(fullAuthorityValue))) {
                  cell.style(STYLE_FILL, appConfig.fullAuthorityColor);
                } else if (partialAuthorityValue.length > 0 && value.match(new RegExp(partialAuthorityValue))) {
                  cell.style(STYLE_FILL, appConfig.partialAuthorityColor);
                } else if (noAuthorityValue.length > 0 && value.match(new RegExp(noAuthorityValue))) {
                  cell.style(STYLE_FILL, appConfig.noAuthorityColor);
                } else if (value.length === 0 && fillsBlankWithNoAuthorityColor) {
                  cell.style(STYLE_FILL, appConfig.noAuthorityColor);
                }
              }
              if (booleanBehavior !== BOOLEAN_NO_OUTPUT) {
                cell.value(convertBoolean(value));
              } else {
                cell.value(value);
              }
              if (orgCnt !== 0) {
                if (value !== beforeValue) {
                  cell = sheet.row(resultWorkY).cell(appConfig.orgDifferentXPosition);
                  cell.value(appConfig.orgDifferentLabel);
                }
                beforeValue = value;
              } else {
                beforeValue = value;
              }
            } else {
              value = null;
              if (baseInfoMap.has(key)) {
                if (notApplicablePermmisonSet && targetName.indexOf(PREFIX_PERMISSION_SET_NAME) === 0) {
                  cell.value(appConfig.notApplicableLabel);
                  cell.style(STYLE_FILL, appConfig.notApplicableColor);
                } else if (settingType === FIELD_LEVEL_SECURITY) {
                  if (fillsBlankWithNoAuthorityColor) {
                    cell.style(STYLE_FILL, appConfig.noAuthorityColor);
                  }
                } else if (booleanBehavior === BOOLEAN_OUTPUT) {
                  cell.value(convertBoolean(METADATA_FALSE));
                  cell.style(STYLE_FILL, appConfig.noAuthorityColor);
                } else if (booleanBehavior === BOOLEAN_OUTPUT_WITH_NA) {
                  cell.value(appConfig.notApplicableLabel);
                  cell.style(STYLE_FILL, appConfig.noAuthorityColor);
                } else if (fillsBlankWithNoAuthorityColor) {
                  cell.style(STYLE_FILL, appConfig.noAuthorityColor);
                }
              } else {
                cell.style(STYLE_FILL, appConfig.notApplicableColor);
              }
              if (orgCnt !== 0) {
                if (value !== beforeValue) {
                  cell = sheet.row(resultWorkY).cell(appConfig.orgDifferentXPosition);
                  cell.value(appConfig.orgDifferentLabel);
                }
                beforeValue = value;
              } else {
                beforeValue = value;
              }  
            }
            resultWorkX++;
          }
        }
        typeWorkY++;
        nameWorkY++;
        secondNameWorkY++;
        labelWorkY++;
        resultWorkY++;
      }
    });

  }
  await workBook.toFileAsync(userConfig.resultFilePath);
  logging('Done.');
})();

function retrieveBaseInfo(metadataList, orgName, isProfile, baseInfoMap) {
  for (const metadata of metadataList) {
    loggingWithTargetName(metadata, isProfile, 'Retrieve base info.');
    const valueMap = new Map();

    if (isProfile) {
      valueMap.set(KEY_BASE_INFO_NAME, metadata.fullName);
    } else {
      valueMap.set(KEY_BASE_INFO_NAME, metadata.label);
    }
    valueMap.set(KEY_BASE_INFO_PERMISSION_SET, isProfile ? METADATA_FALSE : METADATA_TRUE);
    valueMap.set(KEY_BASE_INFO_CUSTOM, metadata.custom);
    if (isProfile) {
      valueMap.set(KEY_BASE_INFO_USER_LICENSE, metadata.userLicense);
    } else {
      valueMap.set(KEY_BASE_INFO_USER_LICENSE, metadata.license);
    }
    valueMap.set(KEY_BASE_INFO_DESCRIPTION, metadata.description);
    const key = getMetadataKey(orgName, isProfile, metadata);
    baseInfoMap.set(key, valueMap);
  }
}

function retrieveObjectPermissions(metadataList, orgName, isProfile, objectPermissionMap) {
  if (!isExecutable(OBJECT_PERMISSION)) {
    return;
  }
  const targetObjectSet = global.targetObjectSet;

  for (const metadata of metadataList) {
    loggingWithTargetName(metadata, isProfile, 'Retrieve object permissions.');

    let objectPermissions = metadata.objectPermissions;
    if (!Array.isArray(objectPermissions)) {
      objectPermissions = [objectPermissions];
    }
    objectPermissions.forEach(function (objectPermission) {
      if (objectPermission === undefined) {
        return;
      }
      if (targetObjectSet.size > 0 &&
        !targetObjectSet.has(objectPermission.object)) {
        return;
      }
      if (!objectPermissionMap.has(objectPermission.object)) {
        objectPermissionMap.set(objectPermission.object, new Map());
      }
      const valueMap = objectPermissionMap.get(objectPermission.object);
      let value = '';
      if (isTrue(objectPermission.allowCreate)) value += appConfig.objectPermissionsLabel.create;
      if (isTrue(objectPermission.allowRead)) value += appConfig.objectPermissionsLabel.read;
      if (isTrue(objectPermission.allowEdit)) value += appConfig.objectPermissionsLabel.edit;
      if (isTrue(objectPermission.allowDelete)) value += appConfig.objectPermissionsLabel.delete;
      if (isTrue(objectPermission.viewAllRecords)) value += appConfig.objectPermissionsLabel.viewAll;
      if (isTrue(objectPermission.modifyAllRecords)) value += appConfig.objectPermissionsLabel.modifyAll;
      const key = getMetadataKey(orgName, isProfile, metadata);
      valueMap.set(key, value);
    });
  }
}

function retrievefieldLevelSecurities(metadataList, orgName, isProfile, fieldLevelSecurityMap, fieldLevelSecurityFieldSet) {
  if (!isExecutable(FIELD_LEVEL_SECURITY)) {
    return;
  }
  const targetObjectSet = global.targetObjectSet;

  for (const metadata of metadataList) {
    loggingWithTargetName(metadata, isProfile, 'Retrieve field-level securities.');
    let fieldPermissions = metadata.fieldPermissions;
    if (!Array.isArray(fieldPermissions)) {
      fieldPermissions = [fieldPermissions];
    }
    fieldPermissions.forEach(function (fieldPermission) {
      if (fieldPermission === undefined) {
        return;
      }
      if (targetObjectSet.size > 0 &&
        !targetObjectSet.has(fieldPermission.field.split('.')[0])) {
        return;
      }
      if (!fieldLevelSecurityMap.has(fieldPermission.field)) {
        fieldLevelSecurityMap.set(fieldPermission.field, new Map());
      }
      fieldLevelSecurityFieldSet.add(fieldPermission.field);
      const valueMap = fieldLevelSecurityMap.get(fieldPermission.field);
      let value = '';
      if (isTrue(fieldPermission.readable)) value += appConfig.fieldLevelSecurityLabel.readable;
      if (isTrue(fieldPermission.editable)) value += appConfig.fieldLevelSecurityLabel.editable;
      const key = getMetadataKey(orgName, isProfile, metadata);
      valueMap.set(key, value);
    });
  }
}

function retrieveLayoutAssignments(metadataList, orgName, isProfile, layoutAssignmentMap) {
  if (!isExecutable(LAYOUT_ASSIGNMENT)) {
    return;
  }
  for (const metadata of metadataList) {
    loggingWithTargetName(metadata, isProfile, 'Retrieve layout assignments.');
    let layoutAssignments = metadata.layoutAssignments;
    if (!Array.isArray(layoutAssignments)) {
      layoutAssignments = [layoutAssignments];
    }
    layoutAssignments.forEach(function (layoutAssignment) {
      if (layoutAssignment === undefined) {
        return;
      }
      let name = layoutAssignment.layout;
      if (layoutAssignment.recordType) {
        const recordTypeRegex = /^.+?\.(.+?)$/.exec(layoutAssignment.recordType);
        name += ' ' + ADD_INFO_HEADER + recordTypeRegex[1];
      } else {
        name += ' ' + ADD_INFO_HEADER + appConfig.layoutAssignLabel.master;
      }

      if (!layoutAssignmentMap.has(name)) {
        layoutAssignmentMap.set(name, new Map());
      }
      const valueMap = layoutAssignmentMap.get(name);
      const key = getMetadataKey(orgName, isProfile, metadata);
      valueMap.set(key, METADATA_TRUE);
    });
  }
}

function retrieveRecordTypeVisibilities(metadataList, orgName, isProfile, recordTypeVisibilityMap) {
  if (!isExecutable(RECORD_TYPE_VISIBILITY)) {
    return;
  }
  for (const metadata of metadataList) {
    loggingWithTargetName(metadata, isProfile, 'Retrieve record-type visibilities.');
    let recordTypeVisibilities = metadata.recordTypeVisibilities;
    if (!Array.isArray(recordTypeVisibilities)) {
      recordTypeVisibilities = [recordTypeVisibilities];
    }
    recordTypeVisibilities.forEach(function (recordTypeVisibility) {
      if (recordTypeVisibility === undefined) {
        return;
      }

      let recordType = recordTypeVisibility.recordType;
      let orgRecordType = recordType;
      recordType = recordType.replace('PersonAccount.', 'Account.');
      if (!recordTypeVisibilityMap.has(recordType)) {
        recordTypeVisibilityMap.set(recordType, new Map());
      }

      const valueMap = recordTypeVisibilityMap.get(recordType);
      const key = getMetadataKey(orgName, isProfile, metadata);
      let isPersonAccountDefault = false;
      let isCompanyAccountDefault = false;
      if ((/^Account\./.exec(recordType) || /^Contact\./.exec(recordType)) &&
        isTrue(recordTypeVisibility.personAccountDefault)) {
        if (/^PersonAccount\./.exec(orgRecordType)) {
          isPersonAccountDefault = true;
        } else {
          isCompanyAccountDefault = true;
        }
      }

      let value = '';
      if (isTrue(recordTypeVisibility.visible)) {
        value += appConfig.recordTypeVisibilityLabel.visible;
      }
      if (isTrue(recordTypeVisibility.default) || isPersonAccountDefault || isCompanyAccountDefault) {
        value += appConfig.recordTypeVisibilityLabel.openBracket;
      }
      if (isTrue(recordTypeVisibility.default)) {
        value += appConfig.recordTypeVisibilityLabel.default;
      }
      if (isPersonAccountDefault) {
        if (isTrue(recordTypeVisibility.default)) {
          value += appConfig.recordTypeVisibilityLabel.delimiter;
        }
        value += appConfig.recordTypeVisibilityLabel.personAccountDefault;
      }
      if (isCompanyAccountDefault) {
        if (isTrue(recordTypeVisibility.default)) {
          value += appConfig.recordTypeVisibilityLabel.delimiter;
        }
        value += appConfig.recordTypeVisibilityLabel.companyAccountDefault;
      }
      if (isTrue(recordTypeVisibility.default) || isPersonAccountDefault || isCompanyAccountDefault) {
        value += appConfig.recordTypeVisibilityLabel.closeBracket;
      }
      valueMap.set(key, value);
    });
  }
}

function retrieveApexClassAccesses(metadataList, orgName, isProfile, apexClassAccessMap) {
  if (!isExecutable(APEX_CLASS_ACCESS)) {
    return;
  }
  for (const metadata of metadataList) {
    loggingWithTargetName(metadata, isProfile, 'Retrieve apex class accesses.');
    let classAccesses = metadata.classAccesses;
    if (!Array.isArray(classAccesses)) {
      classAccesses = [classAccesses];
    }
    classAccesses.forEach(function (classAccess) {
      if (classAccess === undefined) {
        return;
      }
      if (!apexClassAccessMap.has(classAccess.apexClass)) {
        apexClassAccessMap.set(classAccess.apexClass, new Map());
      }
      const valueMap = apexClassAccessMap.get(classAccess.apexClass);
      const key = getMetadataKey(orgName, isProfile, metadata);
      valueMap.set(key, classAccess.enabled);
    });
  }
}

function retrieveApexPageAccesses(metadataList, orgName, isProfile, apexPageAccessMap) {
  if (!isExecutable(APEX_PAGE_ACCESS)) {
    return;
  }
  for (const metadata of metadataList) {
    loggingWithTargetName(metadata, isProfile, 'Retrieve apex page accesses.');
    let pageAccesses = metadata.pageAccesses;
    if (!Array.isArray(pageAccesses)) {
      pageAccesses = [pageAccesses];
    }
    pageAccesses.forEach(function (pageAccess) {
      if (pageAccess === undefined) {
        return;
      }
      if (!apexPageAccessMap.has(pageAccess.apexPage)) {
        apexPageAccessMap.set(pageAccess.apexPage, new Map());
      }
      const valueMap = apexPageAccessMap.get(pageAccess.apexPage);
      const key = getMetadataKey(orgName, isProfile, metadata);
      valueMap.set(key, pageAccess.enabled);
    });
  }
}

function retrieveUserPermissions(metadataList, orgName, isProfile, userPermissionMap) {
  if (!isExecutable(USER_PERMISSION)) {
    return;
  }

  for (const metadata of metadataList) {
    loggingWithTargetName(metadata, isProfile, 'Retrieve user permissions.');
    let userPermissions = metadata.userPermissions;
    if (!Array.isArray(userPermissions)) {
      userPermissions = [userPermissions];
    }
    userPermissions.forEach(function (userPermission) {
      if (userPermission === undefined) {
        return;
      }
      if (!userPermissionMap.has(userPermission.name)) {
        userPermissionMap.set(userPermission.name, new Map());
      }
      const valueMap = userPermissionMap.get(userPermission.name);
      const key = getMetadataKey(orgName, isProfile, metadata);
      valueMap.set(key, userPermission.enabled);
    });
  }
}

function retrieveApplicationVisibilities(metadataList, orgName, isProfile, applicationVisibilityMap) {
  if (!isExecutable(APPLICATION_VISIBILITY)) {
    return;
  }
  const appConfig = global.appConfig;
  for (const metadata of metadataList) {
    loggingWithTargetName(metadata, isProfile, 'Retrieve application visibilities.');
    let applicationVisibilities = metadata.applicationVisibilities;
    if (!Array.isArray(applicationVisibilities)) {
      applicationVisibilities = [applicationVisibilities];
    }
    applicationVisibilities.forEach(function (applicationVisibility) {
      if (applicationVisibility === undefined) {
        return;
      }
      if (!applicationVisibilityMap.has(applicationVisibility.application)) {
        applicationVisibilityMap.set(applicationVisibility.application, new Map());
      }
      const valueMap = applicationVisibilityMap.get(applicationVisibility.application);
      let value = '';
      if (isTrue(applicationVisibility.visible)) {
        value += appConfig.applicationVisibilityLabel.visible;
      }
      if (isTrue(applicationVisibility.default)) {
        value += appConfig.applicationVisibilityLabel.openBracket;
        value += appConfig.applicationVisibilityLabel.default;
        value += appConfig.applicationVisibilityLabel.closeBracket;
      }
      const key = getMetadataKey(orgName, isProfile, metadata);
      valueMap.set(key, value);
    });
  }
}

function retrieveTabVisibilities(metadataList, orgName, isProfile, tabVisibilityMap) {
  if (!isExecutable(TAB_VISIBILITY)) {
    return;
  }
  const appConfig = global.appConfig;
  for (const metadata of metadataList) {
    loggingWithTargetName(metadata, isProfile, 'Retrieve tab visibilities.');
    let tabVisibilities;
    if (isProfile) {
      tabVisibilities = metadata.tabVisibilities;
    } else {
      tabVisibilities = metadata.tabSettings;
    }
    if (!Array.isArray(tabVisibilities)) {
      tabVisibilities = [tabVisibilities];
    }
    tabVisibilities.forEach(function (tabVisibility) {
      if (tabVisibility === undefined) {
        return;
      }
      if (!tabVisibilityMap.has(tabVisibility.tab)) {
        tabVisibilityMap.set(tabVisibility.tab, new Map());
      }
      const valueMap = tabVisibilityMap.get(tabVisibility.tab);
      const key = getMetadataKey(orgName, isProfile, metadata);

      let value = '';
      if (tabVisibility.visibility === 'DefaultOn') {
        value = appConfig.tabVisibilityLabel.defaultOn;
      } else if (tabVisibility.visibility === 'DefaultOff') {
        value = appConfig.tabVisibilityLabel.defaultOff;
      } else if (tabVisibility.visibility === 'Hidden') {
        value = appConfig.tabVisibilityLabel.hidden;
      } else if (tabVisibility.visibility === 'Available') {
        value = appConfig.tabVisibilityLabel.available;
      } else if (tabVisibility.visibility === 'Visible') {
        value = appConfig.tabVisibilityLabel.visible;
      } else {
        value = tabVisibility.visibility;
      }
      valueMap.set(key, value);
    });
  }
}

function retrieveLoginIpRanges(metadataList, orgName, isProfile, loginIpRangeMap) {
  if (!isExecutable(LOGIN_IP_RANGE)) {
    return;
  }
  for (const metadata of metadataList) {
    loggingWithTargetName(metadata, isProfile, 'Retrieve login IP ranges.');
    let loginIpRanges = metadata.loginIpRanges;
    if (!Array.isArray(loginIpRanges)) {
      loginIpRanges = [loginIpRanges];
    }
    loginIpRanges.forEach(function (loginIpRange) {
      if (loginIpRange === undefined) {
        return;
      }
      const ipRange = loginIpRange.startAddress + ' - ' + loginIpRange.endAddress;
      if (!loginIpRangeMap.has(ipRange)) {
        loginIpRangeMap.set(ipRange, new Map());
      }
      const valueMap = loginIpRangeMap.get(ipRange);
      const key = getMetadataKey(orgName, isProfile, metadata);
      valueMap.set(key, METADATA_TRUE);
    });
  }
}

function retrieveCustomPermissions(metadataList, orgName, isProfile, customPermissionMap) {
  if (!isExecutable(CUSTOM_PERMISSION)) {
    return;
  }
  for (const metadata of metadataList) {
    loggingWithTargetName(metadata, isProfile, 'Retrieve custom permissions.');
    let customPermissions = metadata.customPermissions;
    if (!Array.isArray(customPermissions)) {
      customPermissions = [customPermissions];
    }
    customPermissions.forEach(function (customPermission) {
      if (customPermission === undefined) {
        return;
      }
      if (!customPermissionMap.has(customPermission.name)) {
        customPermissionMap.set(customPermission.name, new Map());
      }
      const valueMap = customPermissionMap.get(customPermission.name);
      const key = getMetadataKey(orgName, isProfile, metadata);
      valueMap.set(key, customPermission.enabled);
    });
  }
}

function retrieveSessionSetting(metadataList, orgName, profileName, isProfile, sessionSettingMap) {
  if (!isExecutable(SESSION_SETTING)) {
    return;
  }
  const key = orgName + '.' + PREFIX_PROFILE_NAME + profileName;
  const settingNameArray = [
    KEY_SESSION_SETTING_FORCE_LOGOUT,
    KEY_SESSION_SETTING_REQUIRED_SESSION_LEVEL,
    KEY_SESSION_SETTING_SESSION_PERSISTENCE,
    KEY_SESSION_SETTING_SESSION_TIMEOUT,
    KEY_SESSION_SETTING_SESSION_TIMEOUT_WARNING
  ];
  for (const metadata of metadataList) {
    loggingWithTargetName(metadata, isProfile, 'Retrieve session setting.');

    for (const settingName of settingNameArray) {
      if (!sessionSettingMap.has(settingName)) {
        sessionSettingMap.set(settingName, new Map());
      }
      const valueMap = sessionSettingMap.get(settingName);
      valueMap.set(key, metadata[settingName]);
    }
  }
}

function retrievePasswordPolicy(metadataList, orgName, profileName, isProfile, passwordPolicyMap) {
  if (!isExecutable(PASSWORD_POLICY)) {
    return;
  }
  const key = orgName + '.' + PREFIX_PROFILE_NAME + profileName;
  const settingNameArray = [
    KEY_PASSWORD_POLICY_FORGOT_PASSWORD_REDIRECT,
    KEY_PASSWORD_POLICY_LOCKOUT_INTERVAL,
    KEY_PASSWORD_POLICY_MAX_LOGIN_ATTEMPTS,
    KEY_PASSWORD_POLICY_MINIMUM_PASSWORD_LENGTH,
    KEY_PASSWORD_POLICY_MINIMUM_PASSWORD_LIFE_TIME,
    KEY_PASSWORD_POLICY_OBSCURE,
    KEY_PASSWORD_POLICY_PASSWORD_COMPLEXITY,
    KEY_PASSWORD_POLICY_PASSWORD_EXPIRATION,
    KEY_PASSWORD_POLICY_PASSWORD_HISTORY,
    KEY_PASSWORD_POLICY_PASSWORD_QUESTION
  ];
  for (const metadata of metadataList) {
    loggingWithTargetName(metadata, isProfile, 'Retrieve password policies.');

    for (const settingName of settingNameArray) {
      if (!passwordPolicyMap.has(settingName)) {
        passwordPolicyMap.set(settingName, new Map());
      }
      const valueMap = passwordPolicyMap.get(settingName);
      let value = metadata[settingName];
      if (settingName === KEY_PASSWORD_POLICY_PASSWORD_COMPLEXITY) {
        value = appConfig.passwordComplexityLabel[value];
      }
      valueMap.set(key, value);
    }
  }
}

async function getPermissionSetAPINames(conn) {
  if (global.permissionSetNameList.length === 0) {
    return [];
  }
  let condition = '(';
  for (let i = 0; i < global.permissionSetNameList.length; i++) {
    condition += "'" + global.permissionSetNameList[i] + "'";
    if (i !== global.permissionSetNameList.length - 1) {
      condition += ',';
    }
  }
  condition += ')';

  const recordList = [];
  await conn.query('SELECT Id, Name, Label FROM PermissionSet WHERE Label IN ' + condition)
    .on('record', function (record) { recordList.push(record); })
    .on('error', function (err) { console.error(err); })
    .run({ autoFetch: true });
  const recordMap = new Map();
  for (const i in recordList) {
    recordMap.set(recordList[i].Label, recordList[i].Name);
  }

  const permissionSetAPINameList = [];
  for (const i in global.permissionSetNameList) {
    if (recordMap.has(global.permissionSetNameList[i])) {
      const name = recordMap.get(global.permissionSetNameList[i]);
      permissionSetAPINameList.push(name);
    } else {
      logging('[Warning]Not exist a permission set. Permission set : ' + global.permissionSetNameList[i]);
      continue;
    }
  }

  return permissionSetAPINameList;
}

async function compensateObjectsAndFields(conn, objectPermissionMap, fieldLevelSecurityMap, fieldLevelSecurityFieldSet) {
  if (isExecutable(OBJECT_PERMISSION) || isExecutable(FIELD_LEVEL_SECURITY)) {
    logging('Compensate for object permissions and field-level securities.');

    if (global.targetObjectSet.size === 0) {
      for (const key of objectPermissionMap.keys()) {
        global.targetObjectSet.add(key);
      }
      for (const key of fieldLevelSecurityMap.keys()) {
        global.targetObjectSet.add(key.split('.')[0]);
      }
    }

    for await (const object of global.targetObjectSet) {
      logging('  object: ' + object);
      try {
        await conn.describe(object, function (err, metadata) {
          if (metadata !== undefined) {
            if (err) {
              console.error(err);
              process.exit(1);
            }
            if (isExecutable(OBJECT_PERMISSION)) {
              if (!objectPermissionMap.has(metadata.name)) {
                objectPermissionMap.set(metadata.name, new Map());
              }
              const valueMap = objectPermissionMap.get(metadata.name);
              valueMap.set(LABEL_KEY_NAME, metadata.label);
            }

            if (isExecutable(FIELD_LEVEL_SECURITY)) {
              for (const i in metadata.fields) {
                if (!fieldLevelSecurityFieldSet.has(metadata.name + '.' + metadata.fields[i].name)) {
                  continue;
                }
                if (metadata.name === 'User' && !metadata.fields[i].name.includes('__c')) {
                  continue;
                }
                if (metadata.name === 'Account' && metadata.fields[i].name.includes('__pc')) {
                  continue;
                }
                if (!metadata.fields[i].label) {
                  continue;
                }
                const key = metadata.name + '.' + metadata.fields[i].name;
                if (!fieldLevelSecurityMap.has(key)) {
                  fieldLevelSecurityMap.set(key, new Map());
                }
                const value = fieldLevelSecurityMap.get(key);
                value.set(LABEL_KEY_NAME, metadata.label + '.' + metadata.fields[i].label);
              }
            }
          }
        });
      } catch (e) {
        // some objects raise exceptions, so ignore them
        if (e.message !== 'The requested resource does not exist') {
          throw e;
        }
      }
    }

    // remove unnecessary objects from the retrieved object permissions
    for (const value of global.appConfig.exclusionObjectPermission) {
      objectPermissionMap.delete(value);
    }
  }
}

async function compensateLayoutAssignments(conn, layoutAssignmentMap) {
  if (isExecutable(LAYOUT_ASSIGNMENT)) {
    logging('Compensate for layout assignments.');
    if (global.targetObjectSet.size > 0) {
      for (const key of layoutAssignmentMap.keys()) {
        const object = key.split('-')[0];
        if (!global.targetObjectSet.has(object)) {
          layoutAssignmentMap.delete(key);
        }
      }
    }
  }
}

async function compensateRecordTypeVisibilities(conn, recordTypeVisibilityMap) {
  if (isExecutable(RECORD_TYPE_VISIBILITY)) {
    logging('Compensate for record-type visibilities.');
    const recordList = [];
    await conn.query('SELECT Name, SobjectType, DeveloperName FROM RecordType')
      .on('record', function (record) { recordList.push(record); })
      .on('error', function (err) { console.error(err); })
      .run({ autoFetch: true });
    for (const record of recordList) {
      const key = record.SobjectType + '.' + record.DeveloperName;
      if (!recordTypeVisibilityMap.has(key)) {
        recordTypeVisibilityMap.set(key, new Map());
      }
      const valueMap = recordTypeVisibilityMap.get(key);
      valueMap.set(LABEL_KEY_NAME, record.Name);
    }

    if (global.targetObjectSet.size > 0) {
      for (const key of recordTypeVisibilityMap.keys()) {
        const object = key.split('.')[0];
        if (!global.targetObjectSet.has(object)) {
          recordTypeVisibilityMap.delete(key);
        }
      }
    }
  }
}

async function compensateApexClassAccesses(conn, apexClassAccessMap) {
  if (isExecutable(APEX_CLASS_ACCESS)) {
    logging('Compensate for apex class accesses.');
    const recordList = [];
    await conn.query('SELECT Name FROM ApexClass')
      .on('record', function (record) { recordList.push(record); })
      .on('error', function (err) { console.error(err); })
      .run({ autoFetch: true });
    for (const record of recordList) {
      if (!apexClassAccessMap.has(record.Name)) {
        apexClassAccessMap.set(record.Name, new Map());
      }
    }
  }
}

async function compensateApexPageAccesses(conn, apexPageAccessMap) {
  if (isExecutable(APEX_PAGE_ACCESS)) {
    logging('Compensate for apex page accesses.');
    const recordList = [];
    await conn.query('SELECT Name, MasterLabel FROM ApexPage')
      .on('record', function (record) { recordList.push(record); })
      .on('error', function (err) { console.error(err); })
      .run({ autoFetch: true });
    for (const record of recordList) {
      if (!apexPageAccessMap.has(record.Name)) {
        apexPageAccessMap.set(record.Name, new Map());
      }
      const valueMap = apexPageAccessMap.get(record.Name);
      valueMap.set(LABEL_KEY_NAME, record.MasterLabel);
    }
  }
}

async function compensateUserPermissions(conn, userPermissionMap) {
  if (isExecutable(USER_PERMISSION)) {
    logging('Compensate for user permissons.');
    await conn.sobject('Profile').describe(function (err, metadata) {
      if (metadata !== undefined) {
        if (err) {
          console.error(err);
          process.exit(1);
        }
        for (const i in metadata.fields) {
          let key = metadata.fields[i].name;
          if (!key.indexOf('Permissions')) {
            key = key.replace('Permissions', '');
            if (!userPermissionMap.has(key)) {
              userPermissionMap.set(key, new Map());
            }
            const value = userPermissionMap.get(key);
            value.set(LABEL_KEY_NAME, metadata.fields[i].label);
          }
        }
      }
    });
  }
}

async function compensateApplicationVisibilities(conn, applicationVisibilityMap) {
  if (isExecutable(APPLICATION_VISIBILITY)) {
    logging('Compensate for application visibilities.');
    const recordList = [];
    await conn.query('SELECT NamespacePrefix, DeveloperName, Label FROM AppDefinition')
      .on('record', function (record) { recordList.push(record); })
      .on('error', function (err) { console.error(err); })
      .run({ autoFetch: true });
    for (const record of recordList) {
      let developerName = '';
      if (record.NamespacePrefix) {
        developerName = record.NamespacePrefix + '__' + record.DeveloperName;
      } else {
        developerName = record.DeveloperName;
      }

      if (!applicationVisibilityMap.has(developerName)) {
        applicationVisibilityMap.set(developerName, new Map());
      }
      const valueMap = applicationVisibilityMap.get(developerName);
      valueMap.set(LABEL_KEY_NAME, record.Label);
    }
  }
}

async function compensateTabVisibilities(conn, tabVisibilityMap) {
  if (isExecutable(TAB_VISIBILITY)) {
    logging('Compensate for tab visibilities.');
    const recordList = [];
    await conn.query('SELECT Name, Label FROM TabDefinition')
      .on('record', function (record) { recordList.push(record); })
      .on('error', function (err) { console.error(err); })
      .run({ autoFetch: true });
    for (const record of recordList) {
      if (!tabVisibilityMap.has(record.Name)) {
        tabVisibilityMap.set(record.Name, new Map());
      }
      const valueMap = tabVisibilityMap.get(record.Name);
      valueMap.set(LABEL_KEY_NAME, record.Label);
    }
  }
}

async function compensateCustomPermissions(conn, customPermissionMap) {
  if (isExecutable(CUSTOM_PERMISSION)) {
    logging('Compensate for custom permissions.');
    const recordList = [];
    await conn.query('SELECT NamespacePrefix, DeveloperName, MasterLabel FROM CustomPermission')
      .on('record', function (record) { recordList.push(record); })
      .on('error', function (err) { console.error(err); })
      .run({ autoFetch: true });
    for (const record of recordList) {
      let developerName = '';
      if (record.NamespacePrefix) {
        developerName = record.NamespacePrefix + '__' + record.DeveloperName;
      } else {
        developerName = record.DeveloperName;
      }

      if (!customPermissionMap.has(developerName)) {
        customPermissionMap.set(developerName, new Map());
      }
      const valueMap = customPermissionMap.get(developerName);
      valueMap.set(LABEL_KEY_NAME, record.MasterLabel);
    }
  }
}

async function compensateSessionSetting(conn, sessionSettingMap) {
  if (!isExecutable(SESSION_SETTING)) {
    return;
  }
  for (const key of sessionSettingMap.keys()) {
    const valueMap = sessionSettingMap.get(key);
    const label = appConfig.sessionSettingLabel[key];
    if (label) {
      valueMap.set(LABEL_KEY_NAME, label);
    }
  }
}

async function compensatePasswordPolicy(conn, passwordPolicyMap) {
  if (!isExecutable(PASSWORD_POLICY)) {
    return;
  }
  for (const key of passwordPolicyMap.keys()) {
    const valueMap = passwordPolicyMap.get(key);
    const label = appConfig.passwordPolicyLabel[key];
    if (label) {
      valueMap.set(LABEL_KEY_NAME, label);
    }
  }
}

function getMetadataKey(orgName, isProfile, metadata) {
  let key = '';
  if (isProfile) {
    key = orgName + '.' + PREFIX_PROFILE_NAME + metadata.fullName;
  } else {
    key = orgName + '.' + PREFIX_PERMISSION_SET_NAME + metadata.label;
  }
  return key;
}

function loadYamlFile(filename) {
  const fs = require('fs');
  const yaml = require('js-yaml');
  if (!fs.existsSync(filename)) {
    console.error('File not found. filePath: ' + filename);
    process.exit(1);
  }
  const yamlText = fs.readFileSync(filename, 'utf8');
  return yaml.safeLoad(yamlText);
}

function loadAppConfig() {
  let appConfigPathWork = global.userConfig.appConfigPath;
  if (appConfigPathWork === undefined) {
    appConfigPathWork = DEFAULT_APP_CONFIG_PATH;
  }

  const path = require('path');
  const config = loadYamlFile(path.join(__dirname, appConfigPathWork));
  global.appConfig = config;
}

function loadUserConfig(userConfigPath) {
  let userConfigPathWork = userConfigPath;
  if (userConfigPathWork === undefined) {
    userConfigPathWork = DEFAULT_USER_CONFIG_PATH;
  }
  const path = require('path');
  const config = loadYamlFile(path.join(__dirname, userConfigPathWork));
  global.userConfig = config;

  const targetNameList = [];
  const profileNameList = [];
  const permissionSetNameList = [];
  for (const target of config.target) {
    if (target.ps) {
      permissionSetNameList.push(target.name);
      targetNameList.push(PREFIX_PERMISSION_SET_NAME + target.name);
    } else {
      profileNameList.push(target.name);
      targetNameList.push(PREFIX_PROFILE_NAME + target.name);
    }
  }
  global.targetNameList = targetNameList;
  global.profileNameList = profileNameList;
  global.permissionSetNameList = permissionSetNameList;

  const settingTypeSet = new Set();
  const settingTypeList = config.settingType;
  if (settingTypeList !== undefined) {
    for (const settingType of settingTypeList) {
      settingTypeSet.add(settingType);
    }
  }
  global.settingTypeSet = settingTypeSet;

  const targetObjectSet = new Set();
  const userConfigObject = config.object;
  if (userConfigObject !== undefined) {
    for (const i in userConfigObject) {
      targetObjectSet.add(userConfigObject[i]);
    }
  }
  global.targetObjectSet = targetObjectSet;

  const orgList = [];
  for (const org of config.org) {
    const orgInfo = new _orgInfo(
      org.name,
      org.loginUrl,
      org.apiVersion,
      org.userName,
      org.password
    );
    orgList.push(orgInfo);
  }
  global.orgList = orgList;
}

function _orgInfo(name, loginUrl, apiVersion, userName, password) {
  this.name = name;
  this.loginUrl = loginUrl;
  this.apiVersion = apiVersion;
  this.userName = userName;
  this.password = password;
}

function getTemplateStyleList(sheet) {
  const endColNum = sheet.usedRange().endCell().columnNumber();
  const xlsxTemplateStyleList = [];
  const appConfig = global.appConfig;
  for (let i = 1; i <= endColNum; i++) {
    const style = sheet.cell(appConfig.resultPosition[1], i).style([
      'bold',
      'italic',
      'underline',
      'strikethrough',
      'subscript',
      'superscript',
      'fontSize',
      'fontFamily',
      'fontColor',
      'horizontalAlignment',
      'justifyLastLine',
      'indent',
      'verticalAlignment',
      'wrapText',
      'shrinkToFit',
      'textDirection',
      'textRotation',
      'angleTextCounterclockwise',
      'angleTextClockwise',
      'rotateTextUp',
      'rotateTextDown',
      'verticalText',
      'fill',
      'border',
      'borderColor',
      'borderStyle',
      'diagonalBorderDirection',
      'numberFormat'
    ]);
    xlsxTemplateStyleList.push(style);
  }

  return xlsxTemplateStyleList;
}

function putTemplateStyle(sheet, xlsxTemplateStyleList, cellY) {
  for (let i = 0; i < xlsxTemplateStyleList.length; i++) {
    const cell = sheet.cell(cellY, i + 1);
    cell.style(xlsxTemplateStyleList[i]);
  }
}

function putFirstFrameStyle(sheet, xlsxTemplateStyleList, cellY) {
  if (global.appConfig.boundaryTopBorderOn === true) {
    for (let i = 0; i < xlsxTemplateStyleList.length; i++) {
      const cell = sheet.cell(cellY, i + 1);
      cell.style('topBorder', 'true');
      cell.style('topBorderColor', global.appConfig.boundaryTopBorderColor);
      cell.style('topBorderStyle', global.appConfig.boundaryTopBorderStyle);
    }
  }
}

function usage() {
  console.log('usage: compare-permissions.js [-options]');
  console.log('    -c <pathname> specifies a config file path (default is ./user_config.yaml)');
  console.log("    -s            don't display logs of the execution");
  console.log('    -h            output usage');
  process.exit(0);
}

function convertBoolean(value) {
  let result = '';
  if (value === METADATA_TRUE) {
    result = global.appConfig.trueLabel;
  } else {
    result = global.appConfig.falseLabel;
  }
  return result;
}

function isTrue(value) {
  return value === METADATA_TRUE;
}

function isExecutable(settingType) {
  if (global.settingTypeSet.size === 0) {
    return true;
  }
  if (global.settingTypeSet.has(settingType)) {
    return true;
  }
  return false;
}

function logging(message) {
  _logging(message);
}

function loggingWithTargetName(metadata, isProfile, message) {
  let modifyMessage = '';
  if (isProfile) {
    modifyMessage = '[Profile:' + metadata.fullName + '] ' + message;
  } else {
    modifyMessage = '[PermissionSet:' + metadata.label + '] ' + message;
  }
  _logging(modifyMessage);
}

function _logging(message) {
  if (global.enabledLogging) {
    const nowDate = new Date();
    console.log('[' + getFormattedDateTime(nowDate) + '] ' + message);
  }
}

function getFormattedDateTime(date) {
  let dateString =
    date.getFullYear() + '/' +
    ('0' + (date.getMonth() + 1)).slice(-2) + '/' +
    ('0' + date.getDate()).slice(-2) + ' ' +
    ('0' + date.getHours()).slice(-2) + ':' +
    ('0' + date.getMinutes()).slice(-2) + ':' +
    ('0' + date.getSeconds()).slice(-2);
  return dateString;
}
