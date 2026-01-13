'use strict';

const jsforce = require('jsforce');
const excel = require('exceljs');
const path = require('path');

const METADATA_TRUE = 'true';
const METADATA_FALSE = 'false';
const LABEL_KEY_NAME = '0_Label';
const PREFIX_PROFILE_NAME = '1_';
const PREFIX_PERMISSION_SET_NAME = '2_';
const ADD_INFO_HEADER = 'ADD_INFO:';
const BOOLEAN_OUTPUT = '1';
const BOOLEAN_NO_OUTPUT = '2';
const BOOLEAN_OUTPUT_WITH_NA = '3';
const MAX_EXPLORE_COLS = 255;

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
const CUSTOM_METADATA_TYPE_ACCESS = 'CustomMetadataTypeAccess';
const CUSTOM_SETTING_ACCESS = 'CustomSettingAccess';
const LOGIN_IP_RANGE = 'LoginIpRange';
const LOGIN_HOUR = 'LoginHour';
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

const COMMAND_OPTIONS = {
  [COMMAND_OPTION_SILENT]: () => { global.enabledLogging = false; },
  [COMMAND_OPTION_HELP]: () => { usage(); },
  [COMMAND_OPTION_CONFIG]: (i, args) => {
    if (i + 1 >= args.length) {
      usage();
    }
    userConfigPath = args[i + 1];
    return 1;
  }
};

// analyzes command line options
const params = [];
const args = process.argv.slice(2);
for (let i = 0; i < args.length; i++) {
  const optionHandler = COMMAND_OPTIONS[args[i]];
  if (optionHandler) {
    i += optionHandler(i, args) || 0;
  } else {
    params.push(args[i]);
  }
}

loadUserConfig(userConfigPath);
loadAppConfig();
const userConfig = global.userConfig;
const appConfig = global.appConfig;

log('[Settings]');
log('  AppConfigPath: ' + userConfig.appConfigPath);
log('  TemplateFilePath: ' + userConfig.templateFilePath);
log('  ResultFilePath: ' + userConfig.resultFilePath);
log('  TargetProfiles/PermissionSets: ');
if (userConfig.target) {
  userConfig.target.forEach(value => {
    const target = value.ps ? `${value.name}(PS)` : value.name;
    log(`    ${target}`);
  });
}
log('  TargetSettingTypes: ');
if (userConfig.settingType) {
  userConfig.settingType.forEach(value => {
    log(`    ${value}`);
  });
}
log('  TargetObjects: ');
if (userConfig.object) {
  userConfig.object.forEach(value => {
    log(`    ${value}`);
  });
}

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
  const loginHourMap = new Map();
  const userPermissionMap = new Map();
  const customPermissionMap = new Map();
  const customMetadataTypeAccessMap = new Map();
  const customSettingAccessMap = new Map();
  const layoutAssignmentMap = new Map();
  const sessionSettingMap = new Map();
  const passwordPolicyMap = new Map();

  for (const org of global.orgs) {
    let conn;
    const orgName = org.name;
    log('');
    log('[OrgInfo]');
    log('  Name:' + org.name)

    if (org.clientId && org.clientSecret && org.instanceUrl) {
      log('  InstanceUrl:' + org.instanceUrl)
      log('  ApiVersion:' + org.apiVersion)
      conn = new jsforce.Connection({
        oauth2: { 
          clientId : org.clientId,
          clientSecret : org.clientSecret,
          loginUrl: org.instanceUrl
        },
        version: org.apiVersion
      });
      // authorize with salesforce using OAuth
      const userInfo = await conn.authorize({ grant_type: "client_credentials" })
    } else {
      log('  LoginUrl:' + org.loginUrl)
      log('  ApiVersion:' + org.apiVersion)
      log('  UserName:' + org.userName);
      conn = new jsforce.Connection({ loginUrl: org.loginUrl, version: org.apiVersion });
      // login to salesforce using SOAP API
      await conn.login(org.userName, org.password);
    }

    // retrieve metadata in profiles
    if (global.profileNames.length !== 0) {
      for await (const profileName of global.profileNames) {
        const profileNames = [];
        profileNames.push(profileName);
        try {
            let metadatas = await conn.metadata.read(METADATA_TYPE_PROFILE, profileNames);
            log('[Processing profile: ' + profileName + ']');
            metadatas = [metadatas].flat();
            if (Object.keys(metadatas[0]).length === 0) {
              log('  Unable to find a profile.');
              return;
            }
            retrieveBaseInfo(metadatas, orgName, true, baseInfoMap);
            retrieveObjectPermissions(metadatas, orgName, true, objectPermissionMap);
            retrievefieldLevelSecurities(metadatas, orgName, true, fieldLevelSecurityMap, fieldLevelSecurityFieldSet);
            retrieveLayoutAssignments(metadatas, orgName, true, layoutAssignmentMap);
            retrieveRecordTypeVisibilities(metadatas, orgName, true, recordTypeVisibilityMap);
            retrieveApexClassAccesses(metadatas, orgName, true, apexClassAccessMap);
            retrieveApexPageAccesses(metadatas, orgName, true, apexPageAccessMap);
            retrieveUserPermissions(metadatas, orgName, true, userPermissionMap);
            retrieveApplicationVisibilities(metadatas, orgName, true, applicationVisibilityMap);
            retrieveTabVisibilities(metadatas, orgName, true, tabVisibilityMap);
            retrieveLoginIpRanges(metadatas, orgName, true, loginIpRangeMap);
            retrieveLoginHours(metadatas, orgName, true, loginHourMap);
            retrieveCustomPermissions(metadatas, orgName, true, customPermissionMap);
            retrieveCustomMetadataTypeAccesses(metadatas, orgName, true, customMetadataTypeAccessMap);
            retrieveCustomSettingAccesses(metadatas, orgName, true, customSettingAccessMap);
        } catch (err) {
          console.log(err)
          process.exit(1);
        }
      }
    }

    const profileNameMap = new Map();
    for (const profileName of global.profileNames) {
      const profileLowerName = profileName.toLowerCase(profileName);
      profileNameMap.set(profileLowerName, profileName);
    }

    // retrieve metadata in profile session setting
    if (global.profileNames.length !== 0 && isExecutable(SESSION_SETTING)) {
      const types = [{ type: METADATA_TYPE_SESSION_SETTING, folder: null }];
      const sessionSettings = [];
      try {
        let metadatas = await conn.metadata.list(types, org.apiVersion);
        metadatas = [metadatas].flat().filter(Boolean);
        for (const metadata of metadatas) {
          sessionSettings.push(metadata.fullName);
        }
      } catch (err) {
        console.log(err)
        process.exit(1);
      }

      log('[Processing session setting]');
      for await (const sessionSetting of sessionSettings) {
        try {
          let metadatas = await conn.metadata.read(METADATA_TYPE_SESSION_SETTING, [sessionSetting]);
          metadatas = [metadatas].flat();
          if (Object.keys(metadatas[0]).length === 0) {
            log('  Unable to find a session setting: ' + sessionSetting);
            return;
          }
          const profileName = metadatas[0].profile;
          if (profileNameMap.has(profileName)) {
            log('  profile: ' + profileName + ' sessionSetting: ' + sessionSetting);
            retrieveSessionSetting(metadatas, orgName, profileNameMap.get(profileName), true, sessionSettingMap);
          }
        } catch (err) {
          console.log(err)
          process.exit(1);
        }
      }
    }

    // retrieve metadata in profile password policy
    if (global.profileNames.length !== 0 && isExecutable(PASSWORD_POLICY)) {
      const types = [{ type: METADATA_TYPE_PASSWORD_POLICY, folder: null }];
      const passwordPolicies = [];
      try {
        let metadatas = await conn.metadata.list(types, org.apiVersion);
        metadatas = [metadatas].flat().filter(Boolean);
        for (const metadata of metadatas) {
          passwordPolicies.push(metadata.fullName);
        }
      } catch (err) {
        console.log(err)
        process.exit(1);
      }

      log('[Processing password policy]');
      for await (const passwordPolicy of passwordPolicies) {
        try {
          let metadatas = await conn.metadata.read(METADATA_TYPE_PASSWORD_POLICY, [passwordPolicy]);
          metadatas = [metadatas].flat();
          if (Object.keys(metadatas[0]).length === 0) {
            log('  Unable to find a password policy: ' + passwordPolicy);
            return;
          }
          const profileName = metadatas[0].profile;
          if (profileNameMap.has(profileName)) {
            log('  profile: ' + profileName + ' password policy: ' + passwordPolicy);
            retrievePasswordPolicy(metadatas, orgName, profileNameMap.get(profileName), true, passwordPolicyMap);
          }
        } catch (err) {
          console.log(err)
          process.exit(1);
        }
      }
    }

    if (global.permissionSetNames.length !== 0) {
      // retrieve metadata in permission set
      const permissionSetAPINames = await getPermissionSetAPINames(conn);
      for await (const permissionSetAPIName of permissionSetAPINames) {
        try {
          let metadatas = await conn.metadata.read(METADATA_TYPE_PERMISSION_SET, [permissionSetAPIName]);
          metadatas = [metadatas].flat();

          log('[Processing permission set: ' + permissionSetAPIName + ']');
          retrieveBaseInfo(metadatas, orgName, false, baseInfoMap);
          retrieveObjectPermissions(metadatas, orgName, false, objectPermissionMap);
          retrievefieldLevelSecurities(metadatas, orgName, false, fieldLevelSecurityMap, fieldLevelSecurityFieldSet);
          retrieveLayoutAssignments(metadatas, orgName, false, layoutAssignmentMap);
          retrieveRecordTypeVisibilities(metadatas, orgName, false, recordTypeVisibilityMap);
          retrieveApexClassAccesses(metadatas, orgName, false, apexClassAccessMap);
          retrieveApexPageAccesses(metadatas, orgName, false, apexPageAccessMap);
          retrieveUserPermissions(metadatas, orgName, false, userPermissionMap);
          retrieveApplicationVisibilities(metadatas, orgName, false, applicationVisibilityMap);
          retrieveTabVisibilities(metadatas, orgName, false, tabVisibilityMap);
          retrieveCustomPermissions(metadatas, orgName, false, customPermissionMap);
          retrieveCustomMetadataTypeAccesses(metadatas, orgName, false, customMetadataTypeAccessMap);
          retrieveCustomSettingAccesses(metadatas, orgName, false, customSettingAccessMap);
        } catch (err) {
          console.log(err)
          process.exit(1);
        }
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
    await compensateByEntityDefinition(conn, customMetadataTypeAccessMap, CUSTOM_METADATA_TYPE_ACCESS);
    await compensateByEntityDefinition(conn, customSettingAccessMap, CUSTOM_SETTING_ACCESS);
    await compensateSessionSetting(conn, sessionSettingMap);
    await compensatePasswordPolicy(conn, passwordPolicyMap);

    await conn.logout();

  }

  // export to an excel file
  let workbook = new excel.Workbook();
  const templatePath = resolvePath(userConfig.templateFilePath);
  await workbook.xlsx.readFile(templatePath);

  const styleFill = (cell, color) => {
    cell.style = {...cell.style};
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: color }
    }
  };

  log('');
  log('[Exporting to an Excel file: ' + userConfig.resultFilePath + ']');
  const sheet = workbook.worksheets[0];
  let cell;
  const templateStyles = getTemplateStyles(sheet);

  let targetNameWorkX = appConfig.targetNamePosition[0];
  const targetNameWorkY = appConfig.targetNamePosition[1];
  let targetPermissionSetWorkX = appConfig.targetPermissionSetPosition[0];
  const targetPermissionSetWorkY = appConfig.targetPermissionSetPosition[1];
  let targetLicenseWorkX = appConfig.targetLicensePosition[0];
  const targetLicenseWorkY = appConfig.targetLicensePosition[1];
   let targetCustomWorkX = appConfig.targetCustomPosition[0];
  const targetCustomWorkY = appConfig.targetCustomPosition[1];

  // output base information for profile and permission set
  for (let targetCnt = 0; targetCnt < global.targetNames.length; targetCnt++) {
    let targetName = global.targetNames[targetCnt];
    let color = appConfig.targetGroupColorDefault;
    if (global.orgs.length >= 2) {
      if (targetCnt % 2 === 0) {
        color = appConfig.targetGroupColor1;
      } else {
        color = appConfig.targetGroupColor2;
      }
    }

    for (const org of global.orgs) {
      const key = org.name + '.' + targetName;
      const value = baseInfoMap.get(key);
      if (value) {
        cell = sheet.getCell(targetNameWorkY, targetNameWorkX);
        cell.value = value.get(KEY_BASE_INFO_NAME) + '\n(' + org.name + ')';
        styleFill(cell, color);

        cell = sheet.getCell(targetPermissionSetWorkY, targetPermissionSetWorkX);
        cell.value = convertBoolean(value.get(KEY_BASE_INFO_PERMISSION_SET));
        styleFill(cell, color);

        cell = sheet.getCell(targetLicenseWorkY, targetLicenseWorkX);
        cell.value = value.get(KEY_BASE_INFO_USER_LICENSE);
        styleFill(cell, color);

        cell = sheet.getCell(targetCustomWorkY, targetCustomWorkX);
        cell.value = convertBoolean(value.get(KEY_BASE_INFO_CUSTOM));
        if (isTrue(value.get(KEY_BASE_INFO_PERMISSION_SET))) {
          styleFill(cell, appConfig.notApplicableColor);
        } else {
          styleFill(cell, color);
        }
      } else {
        cell = sheet.getCell(targetNameWorkY, targetNameWorkX);
        cell.value = targetName.slice(2) + '\n(' + org.name + ')';
        styleFill(cell, appConfig.notApplicableColor);

        cell = sheet.getCell(targetPermissionSetWorkY, targetPermissionSetWorkX);
        if (targetName.slice(0, 2) === PREFIX_PERMISSION_SET_NAME) {
          cell.value = convertBoolean(METADATA_TRUE);
        } else {
          cell.value = convertBoolean(METADATA_FALSE);
        }
        styleFill(cell, appConfig.notApplicableColor);

        cell = sheet.getCell(targetLicenseWorkY, targetLicenseWorkX);
        cell.value = '-';
        styleFill(cell, appConfig.notApplicableColor);

        cell = sheet.getCell(targetCustomWorkY, targetCustomWorkX);
        cell.value = '-';
        styleFill(cell, appConfig.notApplicableColor);

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
    const defaultOutputMap = new Map([
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
      [LOGIN_HOUR,
        [LOGIN_HOUR, loginHourMap, BOOLEAN_OUTPUT,
          METADATA_TRUE, '', METADATA_FALSE, false, true]],
      [CUSTOM_PERMISSION,
        [CUSTOM_PERMISSION, customPermissionMap, BOOLEAN_OUTPUT,
          METADATA_TRUE, '', METADATA_FALSE, false, false]],
      [CUSTOM_METADATA_TYPE_ACCESS,
        [CUSTOM_METADATA_TYPE_ACCESS, customMetadataTypeAccessMap, BOOLEAN_OUTPUT,
          METADATA_TRUE, '', METADATA_FALSE, false, false]],
      [CUSTOM_SETTING_ACCESS,
        [CUSTOM_SETTING_ACCESS, customSettingAccessMap, BOOLEAN_OUTPUT,
          METADATA_TRUE, '', METADATA_FALSE, false, false]],
      [SESSION_SETTING,
        [SESSION_SETTING, sessionSettingMap, BOOLEAN_NO_OUTPUT,
          '', '', '', false, true]],
      [PASSWORD_POLICY,
        [PASSWORD_POLICY, passwordPolicyMap, BOOLEAN_NO_OUTPUT,
          '', '', '', false, true]]
    ]);

    // get the type of metadata to output
    const outputs = [];
    if (global.settingTypeSet.size === 0) {
      for (const value of defaultOutputMap.keys()) {
        outputs.push(defaultOutputMap.get(value));
      }
    } else {
      for (const value of global.settingTypeSet) {
        if (!defaultOutputMap.has(value)) {
          log('Incorrect a metadata type name: ' + value);
          process.exit(1);
        }
        outputs.push(defaultOutputMap.get(value));
      }
    }

    const settingTypeLabelMap = global.appConfig.settingTypeLabel;

    outputs.forEach(function (value, index) {
      const settingType = value[0];
      const metadataMap = value[1];
      const sortedMap = new Map([...metadataMap.entries()].sort());
      const booleanBehavior = value[2];
      const fullAuthorityValue = value[3];
      const partialAuthorityValue = value[4];
      const noAuthorityValue = value[5];
      const fillsBlankWithNoAuthorityColor = value[6];
      const notApplicablePermmisonSet = value[7];

      let isFirstRow = true;
      for (const key of sortedMap.keys()) {
        putTemplateStyle(sheet, templateStyles, resultWorkY);
        if (isFirstRow) {
          putFirstFrameStyle(sheet, templateStyles, resultWorkY);
          isFirstRow = false;
        }
        cell = sheet.getCell(typeWorkY, typeWorkX);
        const settingTypeLabel = settingTypeLabelMap[settingType];
        if (settingTypeLabel) {
          cell.value = settingTypeLabel;
        } else {
          cell.value = settingType;
        }

        cell = sheet.getCell(nameWorkY, nameWorkX);
        const nameRegex = new RegExp('^(.+?)( ' + ADD_INFO_HEADER + '(.+?))*$').exec(key);
        cell.value = nameRegex[1];
        cell = sheet.getCell(secondNameWorkY, secondNameWorkX);
        if (nameRegex[3]) {
          cell.value = nameRegex[3];
        }

        const valueMap = metadataMap.get(key);
        if (valueMap.has(LABEL_KEY_NAME)) {
          const value = valueMap.get(LABEL_KEY_NAME);
          cell = sheet.getCell(labelWorkY, labelWorkX);
          cell.value = value;
        }
        resultWorkX = appConfig.resultPosition[0];
        for (const targetName of global.targetNames) {
          let beforeValue;
          let value;
          for (let orgCnt = 0; orgCnt < global.orgs.length; orgCnt++) {
            const org = global.orgs[orgCnt];
            cell = sheet.getCell(resultWorkY, resultWorkX);
            const key = org.name + '.' + targetName;
            if (valueMap.has(key)) {
              value = valueMap.get(key);
              value = String(value);
              if (value !== undefined) {
                if (fullAuthorityValue.length > 0 && value.match(new RegExp(fullAuthorityValue))) {
                  styleFill(cell, appConfig.fullAuthorityColor);

                } else if (partialAuthorityValue.length > 0 && value.match(new RegExp(partialAuthorityValue))) {
                  styleFill(cell, appConfig.partialAuthorityColor);
                } else if (noAuthorityValue.length > 0 && value.match(new RegExp(noAuthorityValue))) {
                  styleFill(cell, appConfig.noAuthorityColor);
                } else if (value.length === 0 && fillsBlankWithNoAuthorityColor) {
                  styleFill(cell, appConfig.noAuthorityColor);
                }
              }
              if (booleanBehavior !== BOOLEAN_NO_OUTPUT) {
                cell.value = convertBoolean(value);
              } else {
                cell.value = value;
              }
              if (orgCnt !== 0) {
                if (value !== beforeValue) {
                  cell = sheet.getCell(resultWorkY, appConfig.orgDifferentXPosition);
                  cell.value = appConfig.orgDifferentLabel;
                }
                beforeValue = value;
              } else {
                beforeValue = value;
              }
            } else {
              value = null;
              if (baseInfoMap.has(key)) {
                if (notApplicablePermmisonSet && targetName.indexOf(PREFIX_PERMISSION_SET_NAME) === 0) {
                  cell.value = appConfig.notApplicableLabel;
                  styleFill(cell, appConfig.notApplicableColor);
                } else if (settingType === FIELD_LEVEL_SECURITY) {
                  if (fillsBlankWithNoAuthorityColor) {
                    styleFill(cell, appConfig.noAuthorityColor);
                  }
                } else if (booleanBehavior === BOOLEAN_OUTPUT) {
                  cell.value = convertBoolean(METADATA_FALSE);
                  styleFill(cell, appConfig.noAuthorityColor);
                } else if (booleanBehavior === BOOLEAN_OUTPUT_WITH_NA) {
                  cell.value = appConfig.notApplicableLabel;
                  styleFill(cell, appConfig.noAuthorityColor);
                } else if (fillsBlankWithNoAuthorityColor) {
                  styleFill(cell, appConfig.noAuthorityColor);
                }
              } else {
                styleFill(cell, appConfig.notApplicableColor);
              }
              if (orgCnt !== 0) {
                if (value !== beforeValue) {
                  cell = sheet.getCell(resultWorkY, appConfig.orgDifferentXPosition);
                  cell.value = appConfig.orgDifferentLabel;
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
  await workbook.xlsx.writeFile(userConfig.resultFilePath);
  log('Done.');
})();

function resolvePath(filePath) {
  return path.isAbsolute(filePath) ? filePath : path.join(__dirname, filePath);
}

function retrieveBaseInfo(metadatas, orgName, isProfile, baseInfoMap) {
  for (const metadata of metadatas) {
    log('  Retrieving base info...');
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

function retrieveObjectPermissions(metadatas, orgName, isProfile, objectPermissionMap) {
  if (!isExecutable(OBJECT_PERMISSION)) {
    return;
  }
  const targetObjectSet = global.targetObjectSet;

  for (const metadata of metadatas) {
    log('  Retrieving object permissions...');

    let objectPermissions = metadata.objectPermissions;
    objectPermissions = [objectPermissions].flat();
    objectPermissions.forEach(function (objectPermission) {
      if (!objectPermission) {
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
      if (isTrue(objectPermission.viewAllFields)) value += appConfig.objectPermissionsLabel.viewAllFields;
      const key = getMetadataKey(orgName, isProfile, metadata);
      valueMap.set(key, value);
    });
  }
}

function retrievefieldLevelSecurities(metadatas, orgName, isProfile, fieldLevelSecurityMap, fieldLevelSecurityFieldSet) {
  if (!isExecutable(FIELD_LEVEL_SECURITY)) {
    return;
  }
  const targetObjectSet = global.targetObjectSet;

  for (const metadata of metadatas) {
    log('  Retrieving field-level security...');
    let fieldPermissions = metadata.fieldPermissions;
    fieldPermissions = [fieldPermissions].flat();
    fieldPermissions.forEach(function (fieldPermission) {
      if (!fieldPermission) {
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

function retrieveLayoutAssignments(metadatas, orgName, isProfile, layoutAssignmentMap) {
  if (!isExecutable(LAYOUT_ASSIGNMENT)) {
    return;
  }
  for (const metadata of metadatas) {
    log('  Retrieving layout assignments...');
    let layoutAssignments = metadata.layoutAssignments;
    layoutAssignments = [layoutAssignments].flat();
    layoutAssignments.forEach(function (layoutAssignment) {
      if (!layoutAssignment) {
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

function retrieveRecordTypeVisibilities(metadatas, orgName, isProfile, recordTypeVisibilityMap) {
  if (!isExecutable(RECORD_TYPE_VISIBILITY)) {
    return;
  }
  for (const metadata of metadatas) {
    log('  Retrieving record-type visibilities...');
    let recordTypeVisibilities = metadata.recordTypeVisibilities;
    recordTypeVisibilities = [recordTypeVisibilities].flat();
    recordTypeVisibilities.forEach(function (recordTypeVisibility) {
      if (!recordTypeVisibility) {
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

function retrieveApexClassAccesses(metadatas, orgName, isProfile, apexClassAccessMap) {
  if (!isExecutable(APEX_CLASS_ACCESS)) {
    return;
  }
  for (const metadata of metadatas) {
    log('  Retrieving apex class accesses...');
    let classAccesses = metadata.classAccesses;
    classAccesses = [classAccesses].flat();
    classAccesses.forEach(function (classAccess) {
      if (!classAccess) {
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

function retrieveApexPageAccesses(metadatas, orgName, isProfile, apexPageAccessMap) {
  if (!isExecutable(APEX_PAGE_ACCESS)) {
    return;
  }
  for (const metadata of metadatas) {
    log('  Retrieving apex page accesses...');
    let pageAccesses = metadata.pageAccesses;
    pageAccesses = [pageAccesses].flat();
    pageAccesses.forEach(function (pageAccess) {
      if (!pageAccess) {
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

function retrieveUserPermissions(metadatas, orgName, isProfile, userPermissionMap) {
  if (!isExecutable(USER_PERMISSION)) {
    return;
  }

  for (const metadata of metadatas) {
    log('  Retrieving user permissions...');
    let userPermissions = metadata.userPermissions;
    userPermissions = [userPermissions].flat();
    userPermissions.forEach(function (userPermission) {
      if (!userPermission) {
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

function retrieveApplicationVisibilities(metadatas, orgName, isProfile, applicationVisibilityMap) {
  if (!isExecutable(APPLICATION_VISIBILITY)) {
    return;
  }
  const appConfig = global.appConfig;
  for (const metadata of metadatas) {
    log('  Retrieving application visibilities...');
    let applicationVisibilities = metadata.applicationVisibilities;
    applicationVisibilities = [applicationVisibilities].flat();
    applicationVisibilities.forEach(function (applicationVisibility) {
      if (!applicationVisibility) {
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

function retrieveTabVisibilities(metadatas, orgName, isProfile, tabVisibilityMap) {
  if (!isExecutable(TAB_VISIBILITY)) {
    return;
  }
  const appConfig = global.appConfig;
  for (const metadata of metadatas) {
    log('  Retrieving tab visibilities...');
    let tabVisibilities;
    if (isProfile) {
      tabVisibilities = metadata.tabVisibilities;
    } else {
      tabVisibilities = metadata.tabSettings;
    }
    tabVisibilities = [tabVisibilities].flat();
    tabVisibilities.forEach(function (tabVisibility) {
      if (!tabVisibility) {
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

function retrieveLoginIpRanges(metadatas, orgName, isProfile, loginIpRangeMap) {
  if (!isExecutable(LOGIN_IP_RANGE)) {
    return;
  }
  for (const metadata of metadatas) {
    log('  Retrieving login IP ranges...');
    let loginIpRanges = metadata.loginIpRanges;
    loginIpRanges = [loginIpRanges].flat();
    loginIpRanges.forEach(function (loginIpRange) {
      if (!loginIpRange) {
        return;
      }
      const ipRange = loginIpRange.startAddress + ' - ' + loginIpRange.endAddress;
      if (!loginIpRangeMap.has(ipRange)) {
        loginIpRangeMap.set(ipRange, new Map());
      }
      const valueMap = loginIpRangeMap.get(ipRange);
      if (loginIpRange.description) {
        valueMap.set(LABEL_KEY_NAME, loginIpRange.description);
      }
      const key = getMetadataKey(orgName, isProfile, metadata);
      valueMap.set(key, METADATA_TRUE);
    });
  }
}

function retrieveLoginHours(metadatas, orgName, isProfile, loginHourMap) {
  if (!isExecutable(LOGIN_HOUR)) {
    return;
  }

  const days = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];
  const convertToTimeFmt = (minutes) => {
    const hours = Math.floor(minutes / 60);
    const mins = minutes % 60;
    return `${hours.toString().padStart(2, '0')}:${mins.toString().padStart(2, '0')}`;
  };
  for (const metadata of metadatas) {
    log('  Retrieving login hours...');
    let loginHours = metadata.loginHours;
    loginHours = [loginHours].flat();
    loginHours.forEach(function (loginHour) {
      if (!loginHour) {
        return;
      }
      const processedHours = days.map((day, index) => {
        const startKey = `${day}Start`;
        const endKey = `${day}End`;

        if (loginHour[startKey] !== undefined && loginHour[endKey] !== undefined) {
            const startTime = convertToTimeFmt(Number(loginHour[startKey]));
            const endTime = convertToTimeFmt(Number(loginHour[endKey]));
            let dayLabel = appConfig.loginHourLabel[day];
            if (!dayLabel) dayLabel = day;
            return `${index + 1}_${dayLabel} ${startTime} - ${endTime}`;
        }
        return null;
      }).filter(Boolean);

      processedHours.forEach((processedHour) => {
        if (!loginHourMap.has(processedHour)) {
          loginHourMap.set(processedHour, new Map());
        }
        const valueMap = loginHourMap.get(processedHour);
        const key = getMetadataKey(orgName, isProfile, metadata);
        valueMap.set(key, METADATA_TRUE);    
      });

    });
  }
}

function retrieveCustomPermissions(metadatas, orgName, isProfile, customPermissionMap) {
  if (!isExecutable(CUSTOM_PERMISSION)) {
    return;
  }
  for (const metadata of metadatas) {
    log('  Retrieving custom permissions...');
    let customPermissions = metadata.customPermissions;
    customPermissions = [customPermissions].flat();
    customPermissions.forEach(function (customPermission) {
      if (!customPermission) {
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

function retrieveCustomMetadataTypeAccesses(metadatas, orgName, isProfile, customMetadataTypeAccessMap) {
  if (!isExecutable(CUSTOM_METADATA_TYPE_ACCESS)) {
    return;
  }
  for (const metadata of metadatas) {
    log('  Retrieving custom metadata type accesses...');
    let customMetadataTypeAccesses = metadata.customMetadataTypeAccesses;
    customMetadataTypeAccesses = [customMetadataTypeAccesses].flat();
    customMetadataTypeAccesses.forEach(function (customMetadataTypeAccess) {
      if (!customMetadataTypeAccess) {
        return;
      }
      if (!customMetadataTypeAccessMap.has(customMetadataTypeAccess.name)) {
        customMetadataTypeAccessMap.set(customMetadataTypeAccess.name, new Map());
      }
      const valueMap = customMetadataTypeAccessMap.get(customMetadataTypeAccess.name);
      const key = getMetadataKey(orgName, isProfile, metadata);
      valueMap.set(key, customMetadataTypeAccess.enabled);
    });
  }
}

function retrieveCustomSettingAccesses(metadatas, orgName, isProfile, customSettingAccessMap) {
  if (!isExecutable(CUSTOM_SETTING_ACCESS)) {
    return;
  }
  for (const metadata of metadatas) {
    log('  Retrieving custom setting accesses...');
    let customSettingAccesses = metadata.customSettingAccesses;
    customSettingAccesses = [customSettingAccesses].flat();
    customSettingAccesses.forEach(function (customSettingAccess) {
      if (!customSettingAccess) {
        return;
      }
      if (!customSettingAccessMap.has(customSettingAccess.name)) {
        customSettingAccessMap.set(customSettingAccess.name, new Map());
      }
      const valueMap = customSettingAccessMap.get(customSettingAccess.name);
      const key = getMetadataKey(orgName, isProfile, metadata);
      valueMap.set(key, customSettingAccess.enabled);
    });
  }
}

function retrieveSessionSetting(metadatas, orgName, profileName, isProfile, sessionSettingMap) {
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
  for (const metadata of metadatas) {
    for (const settingName of settingNameArray) {
      if (!sessionSettingMap.has(settingName)) {
        sessionSettingMap.set(settingName, new Map());
      }
      const valueMap = sessionSettingMap.get(settingName);
      valueMap.set(key, metadata[settingName] ?? '');
    }
  }
}

function retrievePasswordPolicy(metadatas, orgName, profileName, isProfile, passwordPolicyMap) {
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
  for (const metadata of metadatas) {
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
  if (global.permissionSetNames.length === 0) {
    return [];
  }
  let condition = '(';
  for (let i = 0; i < global.permissionSetNames.length; i++) {
    condition += "'" + global.permissionSetNames[i] + "'";
    if (i !== global.permissionSetNames.length - 1) {
      condition += ',';
    }
  }
  condition += ')';

  const records = [];
  await conn.query('SELECT Id, Name, Label, NamespacePrefix FROM PermissionSet WHERE Label IN ' + condition)
    .on('record', function (record) { records.push(record); })
    .on('error', function (err) { console.error(err); process.exit(1); })
    .run({ autoFetch: true });
  const recordMap = new Map();
  for (const i in records) {
    if (records[i].NamespacePrefix) {
      recordMap.set(records[i].Label, records[i].NamespacePrefix + '__' + records[i].Name);
    } else {
      recordMap.set(records[i].Label, records[i].Name);
    }
  }

  const permissionSetAPINames = [];
  for (const i in global.permissionSetNames) {
    if (recordMap.has(global.permissionSetNames[i])) {
      const name = recordMap.get(global.permissionSetNames[i]);
      permissionSetAPINames.push(name);
    } else {
      log('[Warning]Not exist a permission set. Permission set : ' + global.permissionSetNames[i]);
      continue;
    }
  }

  return permissionSetAPINames;
}

async function compensateObjectsAndFields(conn, objectPermissionMap, fieldLevelSecurityMap, fieldLevelSecurityFieldSet) {
  if (isExecutable(OBJECT_PERMISSION) || isExecutable(FIELD_LEVEL_SECURITY)) {
    log('[Compensating object permissions and field-level security]');

    if (global.targetObjectSet.size === 0) {
      for (const key of objectPermissionMap.keys()) {
        global.targetObjectSet.add(key);
      }
      for (const key of fieldLevelSecurityMap.keys()) {
        global.targetObjectSet.add(key.split('.')[0]);
      }
    }

    for await (const object of global.targetObjectSet) {
      log('  object: ' + object);
      try {
        const metadata = await conn.describe(object);
        if (metadata) {
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
      } catch (err) {
        // some objects raise exceptions, so ignore them
        if (err.message !== 'The requested resource does not exist') {
          console.error(err);
          process.exit(1);
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
    log('[Compensating layout assignments]');
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
    log('[Compensating record-type visibilities]');
    const records = [];
    await conn.query('SELECT Name, SobjectType, DeveloperName FROM RecordType')
      .on('record', function (record) { records.push(record); })
      .on('error', function (err) { console.error(err); process.exit(1); })
      .run({ autoFetch: true });
    for (const record of records) {
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
    log('[Compensating apex class accesses]');
    const records = [];
    await conn.query('SELECT Name FROM ApexClass')
      .on('record', function (record) { records.push(record); })
      .on('error', function (err) { console.error(err); process.exit(1); })
      .run({ autoFetch: true });
    for (const record of records) {
      if (!apexClassAccessMap.has(record.Name)) {
        apexClassAccessMap.set(record.Name, new Map());
      }
    }
  }
}

async function compensateApexPageAccesses(conn, apexPageAccessMap) {
  if (isExecutable(APEX_PAGE_ACCESS)) {
    log('[Compensating apex page accesses]');
    const records = [];
    await conn.query('SELECT Name, MasterLabel FROM ApexPage')
      .on('record', function (record) { records.push(record); })
      .on('error', function (err) { console.error(err); process.exit(1); })
      .run({ autoFetch: true });
    for (const record of records) {
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
    log('[Compensating user permissons]');
    try {
      const metadata = await conn.sobject('Profile').describe();
      if (metadata) {
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
    } catch (err) {
      console.error(err);
      process.exit(1);
    }
  }
}

async function compensateApplicationVisibilities(conn, applicationVisibilityMap) {
  if (isExecutable(APPLICATION_VISIBILITY)) {
    log('[Compensating application visibilities]');
    const records = [];
    await conn.query('SELECT NamespacePrefix, DeveloperName, Label FROM AppDefinition')
      .on('record', function (record) { records.push(record); })
      .on('error', function (err) { console.error(err); process.exit(1); })
      .run({ autoFetch: true });
    for (const record of records) {
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
    log('[Compensating tab visibilities]');
    const records = [];
    await conn.query('SELECT Name, Label FROM TabDefinition')
      .on('record', function (record) { records.push(record); })
      .on('error', function (err) { console.error(err); process.exit(1); })
      .run({ autoFetch: true });
    for (const record of records) {
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
    log('[Compensating custom permissions]');
    const records = [];
    await conn.query('SELECT NamespacePrefix, DeveloperName, MasterLabel FROM CustomPermission')
      .on('record', function (record) { records.push(record); })
      .on('error', function (err) { console.error(err); process.exit(1); })
      .run({ autoFetch: true });
    for (const record of records) {
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

async function compensateByEntityDefinition(conn, targetMap, metadataType) {
  if ((metadataType === CUSTOM_METADATA_TYPE_ACCESS && isExecutable(CUSTOM_METADATA_TYPE_ACCESS)) ||
      (metadataType === CUSTOM_SETTING_ACCESS && isExecutable(CUSTOM_SETTING_ACCESS))) {

    if (metadataType === CUSTOM_METADATA_TYPE_ACCESS) log('[Compensating custom metadata type accesses]');
    if (metadataType === CUSTOM_SETTING_ACCESS) log('[Compensating custom setting accesses]');
    const records = [];
    if (targetMap.size === 0) {
      return;
    }
    const qualifiedApiNames 
        = Array.from(targetMap.keys()).map((target) => { return '\'' + target + '\'';}).join(',');
    await conn.tooling.query(
      'SELECT NamespacePrefix, QualifiedApiName, Label FROM EntityDefinition WHERE QualifiedApiName IN (' + qualifiedApiNames + ')')
      .on('record', function (record) { records.push(record); })
      .on('error', function (err) { console.error(err); process.exit(1); })
      .run({ autoFetch: true });
    for (const record of records) {
      let developerName = '';
      if (record.NamespacePrefix) {
        developerName = record.NamespacePrefix + '__' + record.QualifiedApiName;
      } else {
        developerName = record.QualifiedApiName;
      }

      if (!targetMap.has(developerName)) {
        targetMap.set(developerName, new Map());
      }
      const valueMap = targetMap.get(developerName);
      valueMap.set(LABEL_KEY_NAME, record.Label);
    }
  }
}

async function compensateCustomSettingAccesses(conn, customSettingAccessMap) {
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
  return yaml.load(yamlText);
}

function loadAppConfig() {
  let appConfigPathWork = global.userConfig.appConfigPath;
  if (!appConfigPathWork) {
    appConfigPathWork = DEFAULT_APP_CONFIG_PATH;
  }

  const config = loadYamlFile(path.join(__dirname, appConfigPathWork));
  global.appConfig = config;
}

function loadUserConfig(userConfigPath) {
  let userConfigPathWork = userConfigPath;
  if (!userConfigPathWork) {
    userConfigPathWork = DEFAULT_USER_CONFIG_PATH;
  }
  const config = loadYamlFile(path.join(__dirname, userConfigPathWork));
  global.userConfig = config;

  const targetNames = [];
  const profileNames = [];
  const permissionSetNames = [];
  for (const target of config.target) {
    if (target.ps) {
      permissionSetNames.push(target.name);
      targetNames.push(PREFIX_PERMISSION_SET_NAME + target.name);
    } else {
      profileNames.push(target.name);
      targetNames.push(PREFIX_PROFILE_NAME + target.name);
    }
  }
  global.targetNames = targetNames;
  global.profileNames = profileNames;
  global.permissionSetNames = permissionSetNames;

  const settingTypeSet = new Set();
  const settingTypes = config.settingType;
  if (settingTypes) {
    for (const settingType of settingTypes) {
      settingTypeSet.add(settingType);
    }
  }
  global.settingTypeSet = settingTypeSet;

  const targetObjectSet = new Set();
  const userConfigObject = config.object;
  if (userConfigObject) {
    for (const i in userConfigObject) {
      targetObjectSet.add(userConfigObject[i]);
    }
  }
  global.targetObjectSet = targetObjectSet;

  const orgs = [];
  for (const org of config.org) {
    const orgInfo = new _orgInfo(
      org.name,
      org.loginUrl,
      org.apiVersion,
      org.userName,
      org.password,
      org.clientId,
      org.clientSecret,
      org.instanceUrl,
    );
    orgs.push(orgInfo);
  }
  global.orgs = orgs;
}

function _orgInfo(name, loginUrl, apiVersion, userName, password, clientId, clientSecret, instanceUrl) {
  this.name = name;
  this.loginUrl = loginUrl;
  this.apiVersion = apiVersion;
  this.userName = userName;
  this.password = password;
  this.clientId = clientId;
  this.clientSecret = clientSecret;
  this.instanceUrl = instanceUrl;
}

function getTemplateStyles(sheet) {
  const appConfig = global.appConfig;
  const templateStyles = [];
  for (let i = 1; i <= MAX_EXPLORE_COLS; i++) {
    const cell = sheet.getCell(appConfig.resultPosition[1], i);
    if (!cell.font) break;
    const style = {
      font: cell.font,
      alignment: cell.alignment,
      border: cell.border,
      fill: cell.fill,
    }
    templateStyles.push(style);
  }

  return templateStyles;
}

function putTemplateStyle(sheet, templateStyles, cellY) {
  for (let i = 0; i < templateStyles.length; i++) {
    const cell = sheet.getCell(cellY, i + 1);
    cell.font = templateStyles[i].font;
    cell.alignment = templateStyles[i].alignment;
    cell.border = templateStyles[i].border;
    cell.fill = templateStyles[i].fill;
  }
}

function putFirstFrameStyle(sheet, templateStyles, cellY) {
  if (global.appConfig.boundaryTopBorderOn === true) {
    for (let i = 0; i < templateStyles.length; i++) {
      const cell = sheet.getCell(cellY, i + 1);
      let border = {...templateStyles[i].border};
      border.top = {style: global.appConfig.boundaryTopBorderStyle, color: {argb: global.appConfig.boundaryTopBorderColor}};
      cell.border = border;
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
  if (value === METADATA_TRUE || value === true) {
    result = global.appConfig.trueLabel;
  } else {
    result = global.appConfig.falseLabel;
  }
  return result;
}

function isTrue(value) {
  return value === METADATA_TRUE || value === true;
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

function log(message) {
  _log(message);
}

function logWithTargetName(metadata, isProfile, message) {
  let modifyMessage = '';
  if (isProfile) {
    modifyMessage = '[Profile:' + metadata.fullName + '] ' + message;
  } else {
    modifyMessage = '[PermissionSet:' + metadata.label + '] ' + message;
  }
  _log(modifyMessage);
}

function _log(message) {
  if (global.enabledLogging) {
//    const nowDate = new Date();
//    console.log('[' + getFormattedDateTime(nowDate) + '] ' + message);
    console.log(message);
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
