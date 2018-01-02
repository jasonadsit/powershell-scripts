'
' Copyright (c) Microsoft Corporation. All rights reserved.
'
' Windows Software Licensing Management Tool.
'
' Script Name: slmgr.vbs
'

Option Explicit

Dim g_objWMIService, g_strComputer, g_strUserName, g_strPassword
g_strComputer = "."
Dim g_serviceConnected
g_serviceConnected = False

dim g_EchoString
g_EchoString = ""

dim g_objRegistry

Dim g_resourceDictionary, g_resourcesLoaded
Set g_resourceDictionary = CreateObject("Scripting.Dictionary")
g_resourcesLoaded = False

Dim g_DeterminedDisplayFlags
g_DeterminedDisplayFlags = False

Dim g_ShowKmsInfo
Dim g_ShowKmsClientInfo
Dim g_ShowTkaClientInfo
Dim g_ShowTBLInfo
Dim g_ShowPhoneInfo

g_ShowKmsInfo = False
g_ShowKmsClientInfo = false
g_ShowTBLInfo = False
g_ShowPhoneInfo = False

' Messages

'Global options
private const L_optInstallProductKey                  = "ipk"
private const L_optInstallProductKeyUsage             = "Install product key (replaces existing key)"

private const L_optUninstallProductKey                = "upk"
private const L_optUninstallProductKeyUsage           = "Uninstall product key"

private const L_optActivateProduct                    = "ato"
private const L_optActivateProductUsage               = "Activate Windows"

private const L_optDisplayInformation                 = "dli"
private const L_optDisplayInformationUsage            = "Display license information (default: current license)"

private const L_optDisplayInformationVerbose          = "dlv"
private const L_optDisplayInformationUsageVerbose     = "Display detailed license information (default: current license)"

private const L_optExpirationDatime                   = "xpr"
private const L_optExpirationDatimeUsage              = "Expiration date for current license state"

'Advanced options
private const L_optClearPKeyFromRegistry              = "cpky"
private const L_optClearPKeyFromRegistryUsage         = "Clear product key from the registry (prevents disclosure attacks)"

private const L_optInstallLicense                     = "ilc"
private const L_optInstallLicenseUsage                = "Install license"

private const L_optReinstallLicenses                  = "rilc"
private const L_optReinstallLicensesUsage             = "Re-install system license files"

private const L_optDisplayIID                         = "dti"
private const L_optDisplayIIDUsage                    = "Display Installation ID for offline activation"

private const L_optPhoneActivateProduct               = "atp"
private const L_optPhoneActivateProductUsage          = "Activate product with user-provided Confirmation ID"

private const L_optReArmWindows                       = "rearm"
private const L_optReArmWindowsUsage                  = "Reset the licensing status of the machine"

'KMS options

private const L_optSetKmsName                         = "skms"
private const L_optSetKmsNameUsage                    = "Set the name and/or the port for the KMS computer this machine will use. IPv6 address must be specified in the format [hostname]:port"

private const L_optClearKmsName                       = "ckms"
private const L_optClearKmsNameUsage                  = "Clear name of KMS computer used (sets the port to the default)"

private const L_optSetKmsHostCaching                  = "skhc"
private const L_optSetKmsHostCachingUsage             = "Enable KMS host caching"

private const L_optClearKmsHostCaching                = "ckhc"
private const L_optClearKmsHostCachingUsage           = "Disable KMS host caching"

private const L_optSetActivationInterval              = "sai"
private const L_optSetActivationIntervalUsage         = "Set interval (minutes) for unactivated clients to attempt KMS connection. The activation interval must be between 15 minutes (min) and 30 days (max) although the default (2 hours) is recommended."

private const L_optSetRenewalInterval                 = "sri"
private const L_optSetRenewalIntervalUsage            = "Set renewal interval (minutes) for activated clients to attempt KMS connection. The renewal interval must be between 15 minutes (min) and 30 days (max) although the default (7 days) is recommended."

private const L_optSetKmsListenPort                   = "sprt"
private const L_optSetKmsListenPortUsage              = "Set TCP port KMS will use to communicate with clients"

private const L_optSetDNS                             = "sdns"
private const L_optSetDNSUsage                        = "Enable DNS publishing by KMS (default)"

private const L_optClearDNS                           = "cdns"
private const L_optClearDNSUsage                      = "Disable DNS publishing by KMS"

private const L_optSetNormalPriority                  = "spri"
private const L_optSetNormalPriorityUsage             = "Set KMS priority to normal (default)"

private const L_optClearNormalPriority                = "cpri"
private const L_optClearNormalPriorityUsage           = "Set KMS priority to low"

' Token-based Activation options

private const L_optListInstalledILs                   = "lil"
private const L_optListInstalledILsUsage              = "List installed Token-based Activation Issuance Licenses"

private const L_optRemoveInstalledIL                  = "ril"
private const L_optRemoveInstalledILUsage             = "Remove installed Token-based Activation Issuance License"

private const L_optClearTkaOnly                       = "ctao"
private const L_optClearTkaOnlyUsage                  = "Clear Token-based Activation Only flag (default)"

private const L_optSetTkaOnly                         = "stao"
private const L_optSetTkaOnlyUsage                    = "Set Token-based Activation Only flag"

private const L_optListTkaCerts                       = "ltc"
private const L_optListTkaCertsUsage                  = "List Token-based Activation Certificates"

private const L_optForceTkaActivation                 = "fta"
private const L_optForceTkaActivationUsage            = "Force Token-based Activation"

' Option parameters
private const L_ParamsActivationID                    = "<Activation ID>"
private const L_ParamsActivationIDOptional            = "[Activation ID]"
private const L_ParamsActIDOptional                   = "[Activation ID | All]"
private const L_ParamsProductKey                      = "<Product Key>"
private const L_ParamsLicenseFile                     = "<License file>"
private const L_ParamsPhoneActivate                   = "<Confirmation ID>"
private const L_ParamsSetKms                          = "<Name[:Port] | : port> [Activation ID]"
private const L_ParamsSetListenKmsPort                = "<Port>"
private const L_ParamsSetActivationInterval           = "<Activation Interval>"
private const L_ParamsSetRenewalInterval              = "<Renewal Interval>"

private const L_ParamsRemoveInstalledIL               = "<ILID> <ILvID>"
private const L_ParamsForceTkaActivation              = "<Certificate Thumbprint> [<PIN>]"

' Miscellaneous messages
private const L_MsgHelp_1                             = "Windows Software Licensing Management Tool"
private const L_MsgHelp_2                             = "Usage: slmgr.vbs [MachineName [User Password]] [<Option>]"
private const L_MsgHelp_3                             = "MachineName: Name of remote machine (default is local machine)"
private const L_MsgHelp_4                             = "User:        Account with required privilege on remote machine"
private const L_MsgHelp_5                             = "Password:    password for the previous account"
private const L_MsgGlobalOptions                      = "Global Options:"
private const L_MsgAdvancedOptions                    = "Advanced Options:"
private const L_MsgKmsClientOptions                   = "Volume Licensing: Key Management Service (KMS) Client Options:"
private const L_MsgKmsOptions                         = "Volume Licensing: Key Management Service (KMS) Options:"
private const L_MsgTkaClientOptions                   = "Volume Licensing: Token-based Activation Options:"
private const L_MsgInvalidOptions                     = "Invalid combination of command parameters."
private const L_MsgUnrecognizedOption                 = "Unrecognized option: "
private const L_MsgErrorProductNotFound               = "Error: product not found."
private const L_MsgClearedPKey                        = "Product key from registry cleared successfully."
private const L_MsgInstalledPKey                      = "Installed product key %PKEY% successfully."
private const L_MsgUninstalledPKey                    = "Uninstalled product key successfully."
private const L_MsgErrorPKey                          = "Error: product key not found."
private const L_MsgInstallationID                     = "Installation ID: "
private const L_MsgPhoneNumbers                       = "Product activation telephone numbers can be obtained by searching the phone.inf file for the appropriate phone number for your location/country. You can open the phone.inf file from a Command Prompt or the Start Menu by running: notepad %systemroot%\system32\sppui\phone.inf"
private const L_MsgActivating                         = "Activating %PRODUCTNAME% (%PRODUCTID%) ..."
private const L_MsgActivated                          = "Product activated successfully."
private const L_MsgActivated_Failed                   = "Error: Product activation failed."
private const L_MsgConfID                             = "Confirmation ID for product %ACTID% deposited successfully."
private const L_MsgErrorDescription                   = "Error description: "
private const L_MsgErrorConnection                    = "Error 0x%ERRCODE% occurred in connecting to server %COMPUTERNAME%."
private const L_MsgInfoRemoteConnection               = "Connected to server %COMPUTERNAME%."
private const L_MsgErrorConnectionRegistry            = "Error 0x%ERRCODE% occurred in connecting to the registry on server %COMPUTERNAME%."
private const L_MsgErrorImpersonation                 = "Error 0x%ERRCODE% occurred in setting impersonation level."
private const L_MsgErrorAuthenticationLevel           = "Error 0x%ERRCODE% occurred in setting authentication level."
private const L_MsgErrorWMI                           = "Error 0x%ERRCODE% occurred in creating a locator object."
private const L_MsgErrorText_6                        = "On a computer running Microsoft Windows non-core edition, run 'slui.exe 0x2a 0x%ERRCODE%' to display the error text."
private const L_MsgErrorText_8                        = "Error: "
private const L_MsgErrorText_9                        = "Error: option %OPTION% needs %PARAM%"
private const L_MsgErrorText_11                       = "The machine is running within the non-genuine grace period. Run 'slui.exe' to go online and make the machine genuine."
private const L_MsgErrorText_12                       = "Windows is running within the non-genuine notification period. Run 'slui.exe' to go online and validate Windows."
private const L_MsgLicenseFile                        = "License file %LICENSEFILE% installed successfully."
private const L_MsgKmsPriSetToLow                     = "KMS priority set to Low"
private const L_MsgKmsPriSetToNormal                  = "KMS priority set to Normal"
private const L_MsgWarningKmsPri                      = "Warning: Priority can only be set on a KMS machine that is also activated."
private const L_MsgKmsDnsPublishingDisabled           = "DNS publishing disabled"
private const L_MsgKmsDnsPublishingEnabled            = "DNS publishing enabled"
private const L_MsgKmsDnsPublishingWarning            = "Warning: DNS Publishing can only be set on a KMS machine that is also activated."
private const L_MsgKmsPortSet                         = "KMS port set to %PORT% successfully."
private const L_MsgWarningKmsReboot                   = "Warning: a KMS reboot is needed for this setting to take effect."
private const L_MsgWarningKmsPort                     = "Warning: KMS port can only be set on a KMS machine that is also activated."
private const L_MsgRenewalSet                         = "Volume renewal interval set to %RENEWAL% minutes successfully."
private const L_MsgWarningRenewal                     = "Warning: Volume renewal interval can only be set on a KMS machine that is also activated."
private const L_MsgActivationSet                      = "Volume activation interval set to %ACTIVATION% minutes successfully."
private const L_MsgWarningActivation                  = "Warning: Volume activation interval can only be set on a KMS machine that is also activated."
private const L_MsgKmsNameSet                         = "Key Management Service machine name set to %KMS% successfully."
private const L_MsgKmsNameCleared                     = "Key Management Service machine name cleared successfully."
private const L_MsgKmsHostCachingDisabled             = "KMS host caching is disabled"
private const L_MsgKmsHostCachingEnabled              = "KMS host caching is enabled"
private const L_MsgErrorActivationID                  = "Error: Activation ID (%ActID%) not found."
private const L_MsgRearm_1                            = "Command completed successfully."
private const L_MsgRearm_2                            = "Please restart the system for the changes to take effect."
private const L_MsgRemainingWindowsRearmCount         = "Remaining Windows rearm count: %COUNT%"
' Used for xpr
private const L_MsgLicenseStatusUnlicensed            = "Unlicensed"
private const L_MsgLicenseStatusVL                    = "Volume activation will expire %ENDDATE%"
private const L_MsgLicenseStatusTBL                   = "Timebased activation will expire %ENDDATE%"
private const L_MsgLicenseStatusLicensed              = "The machine is permanently activated."
private const L_MsgLicenseStatusInitialGrace          = "Initial grace period ends %ENDDATE%"
private const L_MsgLicenseStatusAdditionalGrace       = "Additional grace period ends %ENDDATE%"
private const L_MsgLicenseStatusNonGenuineGrace       = "Non-genuine grace period ends %ENDDATE%"
private const L_MsgLicenseStatusNotification          = "Windows is in Notification mode"
private const L_MsgLicenseStatusExtendedGrace         = "Extended grace period ends %ENDDATE%"

' Used for dli/dlv
private const L_MsgLicenseStatusUnlicensed_1          = "License Status: Unlicensed"
private const L_MsgLicenseStatusLicensed_1            = "License Status: Licensed"
private const L_MsgLicenseStatusVL_1                  = "Volume activation expiration: %MINUTE% minute(s) (%DAY% day(s))"
private const L_MsgLicenseStatusTBL_1                 = "Timebased activation expiration: %MINUTE% minute(s) (%DAY% day(s))"
private const L_MsgLicenseStatusInitialGrace_1        = "License Status: Initial grace period"
private const L_MsgLicenseStatusAdditionalGrace_1     = "License Status: Additional grace period (KMS license expired or hardware out of tolerance)"
private const L_MsgLicenseStatusNonGenuineGrace_1     = "License Status: Non-genuine grace period."
private const L_MsgLicenseStatusNotification_1        = "License Status: Notification"
private const L_MsgLicenseStatusExtendedGrace_1       = "License Status: Extended grace period"

private const L_MsgNotificationErrorReasonNonGenuine  = "Notification Reason: 0x%ERRCODE% (non-genuine)."
private const L_MsgNotificationErrorReasonExpiration  = "Notification Reason: 0x%ERRCODE% (grace time expired)."
private const L_MsgNotificationErrorReasonOther       = "Notification Reason: 0x%ERRCODE%."
private const L_MsgLicenseStatusTimeRemaining         = "Time remaining: %MINUTE% minute(s) (%DAY% day(s))"
private const L_MsgLicenseStatusUnknown               = "License Status: Unknown"
private const L_MsgLicenseStatusEvalEndData           = "Evaluation End Date: "
private const L_MsgReinstallingLicenses               = "Re-installing license files ..."
private const L_MsgLicensesReinstalled                = "License files re-installed successfully."
private const L_MsgServiceVersion                     = "Software licensing service version: "
private const L_MsgProductName                        = "Name: "
private const L_MsgProductDesc                        = "Description: "
private const L_MsgActID                              = "Activation ID: "
private const L_MsgAppID                              = "Application ID: "
private const L_MsgPID4                               = "Extended PID: "
private const L_MsgProcessorCertUrl                   = "Processor Certificate URL: "
private const L_MsgMachineCertUrl                     = "Machine Certificate URL: "
private const L_MsgUseLicenseCertUrl                  = "Use License URL: "
private const L_MsgPKeyCertUrl                        = "Product Key Certificate URL: "
private const L_MsgPartialPKey                        = "Partial Product Key: "
private const L_MsgErrorLicenseNotInUse               = "This license is not in use."
private const L_MsgKmsInfo                            = "Key Management Service client information"
private const L_MsgCmid                               = "Client Machine ID (CMID): "
private const L_MsgRegisteredKmsName                  = "Registered KMS machine name: "
private const L_MsgKmsFromDnsUnavailable              = "DNS auto-discovery: KMS name not available"
private const L_MsgKmsFromDns                         = "KMS machine name from DNS: "
private const L_MsgKmsPID4                            = "KMS machine extended PID: "
private const L_MsgActivationInterval                 = "Activation interval: %INTERVAL% minutes"
private const L_MsgRenewalInterval                    = "Renewal interval: %INTERVAL% minutes"
private const L_MsgKmsEnabled                         = "Key Management Service is enabled on this machine"
private const L_MsgKmsCurrentCount                    = "Current count: "
private const L_MsgKmsListeningOnPort                 = "Listening on Port: "
private const L_MsgKmsPriNormal                       = "KMS priority: Normal"
private const L_MsgKmsPriLow                          = "KMS priority: Low"
private const L_MsgInvalidDataError                   = "Error: The data is invalid"
private const L_MsgUndeterminedPrimaryKey             = "Warning: SLMGR was not able to validate the current product key for Windows. Please upgrade to the latest service pack."
private const L_MsgUndeterminedPrimaryKeyOperation    = "Warning: This operation may affect more than one target license.  Please verify the results."
private const L_MsgUndeterminedOperationFormat        = "Processing the license for %PRODUCTDESCRIPTION% (%PRODUCTID%)."
private const L_MsgPleaseActivateRefreshKMSInfo       = "Please use slmgr.vbs /ato to activate and update KMS client information in order to update values."

private const L_MsgKmsCumulativeRequestsFromClients             = "Key Management Service cumulative requests received from clients"
private const L_MsgKmsTotalRequestsRecieved                     = "Total requests received: "
private const L_MsgKmsFailedRequestsReceived                    = "Failed requests received: "
private const L_MsgKmsRequestsWithStatusUnlicensed              = "Requests with License Status Unlicensed: "
private const L_MsgKmsRequestsWithStatusLicensed                = "Requests with License Status Licensed: "
private const L_MsgKmsRequestsWithStatusInitialGrace            = "Requests with License Status Initial grace period: "
private const L_MsgKmsRequestsWithStatusLicenseExpiredOrHwidOot = "Requests with License Status License expired or Hardware out of tolerance: "
private const L_MsgKmsRequestsWithStatusNonGenuineGrace         = "Requests with License Status Non-genuine grace period: "
private const L_MsgKmsRequestsWithStatusNotification            = "Requests with License Status Notification: "

private const L_MsgRemoteWmiVersionMismatch           = "The remote machine does not support this version of SLMgr.vbs"

'
' Token-based Activation issuance licenses
'
private const L_MsgTkaLicenses                        = "Token-based Activation Issuance Licenses:"
private const L_MsgTkaLicenseHeader                   = "%ILID%    %ILVID%"
private const L_MsgTkaLicenseILID                     = "License ID (ILID): %ILID%"
private const L_MsgTkaLicenseILVID                    = "Version ID (ILvID): %ILVID%"
private const L_MsgTkaLicenseExpiration               = "Valid to: %TODATE%"
private const L_MsgTkaLicenseAdditionalInfo           = "Additional Information: %MOREINFO%"
private const L_MsgTkaLicenseAuthZStatus              = "Error: 0x%ERRCODE%"
private const L_MsgTkaLicenseDescr                    = "Description: %DESC%"
private const L_MsgTkaLicenseNone                     = "No licenses found."

private const L_MsgTkaRemoving                        = "Removing Token-based Activation License ..."
private const L_MsgTkaRemovedItem                     = "Removed license with SLID=%SLID%."
private const L_MsgTkaRemovedNone                     = "No licenses found."

private const L_MsgTkaInfoAdditionalInfo              = "Additional Information: %MOREINFO%" 
private const L_MsgTkaInfo                            = "Token-based Activation information"
private const L_MsgTkaInfoILID                        = "License ID (ILID): %ILID%"
private const L_MsgTkaInfoILVID                       = "Version ID (ILvID): %ILVID%"
private const L_MsgTkaInfoGrantNo                     = "Grant Number: %GRANTNO%"
private const L_MsgTkaInfoThumbprint                  = "Certificate Thumbprint: %THUMBPRINT%"

private const L_MsgTkaTaoSet                          = "The machine will activate via Token-based Activation only"
private const L_MsgTkaTaoClear                        = "The machine will activate via KMS and Token-based Activation"

private const L_MsgTkaCertThumbprint                  = "Thumbprint: %THUMBPRINT%"
private const L_MsgTkaCertSubject                     = "Subject: %SUBJECT%"
private const L_MsgTkaCertIssuer                      = "Issuer: %ISSUER%"
private const L_MsgTkaCertValidFrom                   = "Valid from: %FROMDATE%"
private const L_MsgTkaCertValidTo                     = "Valid to: %TODATE%"

private const L_MsgCurrentTrustedTime                 = "Trusted time: "

private const NoPrimaryKeyFound                       = "NoPrimaryKeyFound"
private const TblPrimaryKey                           = "TblPrimaryKey"
private const NotSpecialCasePrimaryKey                = "NotSpecialCasePrimaryKey"
private const IndeterminatePrimaryKeyFound            = "IndeterminatePrimaryKey"

private const L_MsgError_C004C001                     = "The activation server determined the specified product key is invalid"
private const L_MsgError_C004C003                     = "The activation server determined the specified product key is blocked"
private const L_MsgError_C004C017                     = "The activation server determined the specified product key has been blocked for this geographic location."
private const L_MsgError_C004B100                     = "The activation server determined that the computer could not be activated"
private const L_MsgError_C004C008                     = "The activation server determined that the specified product key could not be used"
private const L_MsgError_C004C020                     = "The activation server reported that the Multiple Activation Key has exceeded its limit"
private const L_MsgError_C004C021                     = "The activation server reported that the Multiple Activation Key extension limit has been exceeded"
private const L_MsgError_C004D307                     = "The maximum allowed number of re-arms has been exceeded. You must re-install the OS before trying to re-arm again"
private const L_MsgError_C004F009                     = "The software Licensing Service reported that the grace period expired"
private const L_MsgError_C004F00F                     = "The Software Licensing Server reported that the hardware ID binding is beyond level of tolerance"
private const L_MsgError_C004F014                     = "The Software Licensing Service reported that the product key is not available"
private const L_MsgError_C004F025                     = "Access denied: the requested action requires elevated privileges"
private const L_MsgError_C004F02C                     = "The software Licensing Service reported that the format for the offline activation data is incorrect"
private const L_MsgError_C004F035                     = "The software Licensing Service reported that the computer could not be activated with a Volume license product key. Volume licensed systems require upgrading from a qualified operating system. Please contact your system administrator or use a different type of key"
private const L_MsgError_C004F038                     = "The software Licensing Service reported that the computer could not be activated. The count reported by your Key Management Service (KMS) is insufficient. Please contact your system administrator"
private const L_MsgError_C004F039                     = "The software Licensing Service reported that the computer could not be activated. The Key Management Service (KMS) is not enabled"
private const L_MsgError_C004F041                     = "The software Licensing Service determined that the Key Management Server (KMS) is not activated. KMS needs to be activated"
private const L_MsgError_C004F042                     = "The software Licensing Service determined that the specified Key Management Service (KMS) cannot be used"
private const L_MsgError_C004F050                     = "The Software Licensing Service reported that the product key is invalid"
private const L_MsgError_C004F051                     = "The software Licensing Service reported that the product key is blocked"
private const L_MsgError_C004F064                     = "The software Licensing Service reported that the non-Genuine grace period expired"
private const L_MsgError_C004F065                     = "The software Licensing Service reported that the application is running within the valid non-genuine period"
private const L_MsgError_C004F066                     = "The Software Licensing Service reported that the product SKU is not found"
private const L_MsgError_C004F06B                     = "The software Licensing Service determined that it is running in a virtual machine. The Key Management Service (KMS) is not supported in this mode"
private const L_MsgError_C004F074                     = "The Software Licensing Service reported that the computer could not be activated. The Key Management Service (KMS) is unavailable"
private const L_MsgError_C004F075                     = "The Software Licensing Service reported that the operation cannot be completed because the service is stopping"

private const L_MsgError_C004F304                     = "The Software Licensing Service reported that required license could not be found."
private const L_MsgError_C004F305                     = "The Software Licensing Service reported that there are no certificates found in the system that could activate the product."
private const L_MsgError_C004F30A                     = "The Software Licensing Service reported that the computer could not be activated. The certificate does not match the conditions in the license."
private const L_MsgError_C004F30D                     = "The Software Licensing Service reported that the computer could not be activated. The thumbprint is invalid."
private const L_MsgError_C004F30E                     = "The Software Licensing Service reported that the computer could not be activated. A certificate for the thumbprint could not be found."

private const L_MsgError_C004F30F                     = "The Software Licensing Service reported that the computer could not be activated. The certificate does not match the criteria specified in the issuance license."
private const L_MsgError_C004F310                     = "The Software Licensing Service reported that the computer could not be activated. The certificate does not match the trust point identifier (TPID) specified in the issuance license."
private const L_MsgError_C004F311                     = "The Software Licensing Service reported that the computer could not be activated. A soft token cannot be used for activation."
private const L_MsgError_C004F312                     = "The Software Licensing Service reported that the computer could not be activated. The certificate cannot be used because its private key is exportable."

private const L_MsgError_5                            = "Access denied: the requested action requires elevated privileges"
private const L_MsgError_80070005                     = "Access denied: the requested action requires elevated privileges"
private const L_MsgError_80070057                     = "The parameter is incorrect"
private const L_MsgError_8007232A                     = "DNS server failure"
private const L_MsgError_8007232B                     = "DNS name does not exist"
private const L_MsgError_800706BA                     = "The RPC server is unavailable"
private const L_MsgError_8007251D                     = "No records found for DNS query"

' Registry constants
private const HKEY_LOCAL_MACHINE                      = &H80000002
private const HKEY_NETWORK_SERVICE                    = &H80000003

private const DefaultPort                             = "1688"
private const intKnownOption                          = 0
private const intUnknownOption                        = 1

private const SLKeyPath                               = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
private const SLKeyPath32                             = "SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
private const NSKeyPath                               = "S-1-5-20\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"

private const HR_S_OK                                 = 0
private const HR_ERROR_FILE_NOT_FOUND                 = &H80070002
private const HR_SL_E_GRACE_TIME_EXPIRED              = &HC004F009
private const HR_SL_E_NOT_GENUINE                     = &HC004F200

' WMI class names
private const ServiceClass                            = "SoftwareLicensingService"
private const ProductClass                            = "SoftwareLicensingProduct"
private const TkaLicenseClass                         = "SoftwareLicensingTokenActivationLicense"
private const WindowsAppId                            = "55c92734-d682-4d71-983e-d6ec3f16059f"

private const ProductIsPrimarySkuSelectClause         = "ID, ApplicationId, PartialProductKey, LicenseIsAddon, Description, Name"

private const PartialProductKeyNonNullWhereClause     = "PartialProductKey <> null"
private const EmptyWhereClause                        = ""

private const wbemImpersonationLevelImpersonate       = 3
private const wbemAuthenticationLevelPktPrivacy       = 6

Call ExecCommandLine()
ExitScript 0

Private Sub DisplayUsage ()

    LineOut GetResource("L_MsgHelp_1")
    LineOut GetResource("L_MsgHelp_2")
    LineOut "           " & GetResource("L_MsgHelp_3")
    LineOut "           " & GetResource("L_MsgHelp_4")
    LineOut "           " & GetResource("L_MsgHelp_5")
    LineOut ""
    LineOut GetResource("L_MsgGlobalOptions")
    OptLine GetResource("L_optInstallProductKey"),         GetResource("L_ParamsProductKey"),            GetResource("L_optInstallProductKeyUsage")
    OptLine GetResource("L_optActivateProduct"),           GetResource("L_ParamsActivationIDOptional"),  GetResource("L_optActivateProductUsage")
    OptLine GetResource("L_optDisplayInformation"),        GetResource("L_ParamsActIDOptional"),         GetResource("L_optDisplayInformationUsage")
    OptLine GetResource("L_optDisplayInformationVerbose"), GetResource("L_ParamsActIDOptional"),         GetResource("L_optDisplayInformationUsageVerbose")
    OptLine GetResource("L_optExpirationDatime"),          GetResource("L_ParamsActivationIDOptional"),  GetResource("L_optExpirationDatimeUsage")

    LineFlush ""
    
    LineOut GetResource("L_MsgAdvancedOptions")
    OptLine GetResource("L_optClearPKeyFromRegistry"),     "",                                           GetResource("L_optClearPKeyFromRegistryUsage")
    OptLine GetResource("L_optInstallLicense"),            GetResource("L_ParamsLicenseFile"),           GetResource("L_optInstallLicenseUsage")
    OptLine GetResource("L_optReinstallLicenses"),         "",                                           GetResource("L_optReinstallLicensesUsage")
    OptLine GetResource("L_optReArmWindows"),              "",                                           GetResource("L_optReArmWindowsUsage")
    OptLine GetResource("L_optUninstallProductKey"),       GetResource("L_ParamsActivationIDOptional"),  GetResource("L_optUninstallProductKeyUsage")


    LineOut ""
    OptLine  GetResource("L_optDisplayIID"),           GetResource("L_ParamsActivationIDOptional"),  GetResource("L_optDisplayIIDUsage")
    OptLine2 GetResource("L_optPhoneActivateProduct"), GetResource("L_ParamsPhoneActivate"),         GetResource("L_ParamsActivationIDOptional"),   GetResource("L_optPhoneActivateProductUsage")

    LineOut ""
    LineOut  GetResource("L_MsgKmsClientOptions")
    OptLine2 GetResource("L_optSetKmsName"),           GetResource("L_ParamsSetKms"),                GetResource("L_ParamsActivationIDOptional"),   GetResource("L_optSetKmsNameUsage")
    OptLine  GetResource("L_optClearKmsName"),         GetResource("L_ParamsActivationIDOptional"),  GetResource("L_optClearKmsNameUsage")
    OptLine GetResource("L_optSetKmsHostCaching"),     "",                                           GetResource("L_optSetKmsHostCachingUsage")
    OptLine GetResource("L_optClearKmsHostCaching"),   "",                                           GetResource("L_optClearKmsHostCachingUsage")        

    LineFlush ""

    LineOut GetResource("L_MsgTkaClientOptions")
    OptLine GetResource("L_optListInstalledILs"),      "",                                           GetResource("L_optListInstalledILsUsage")
    OptLine GetResource("L_optRemoveInstalledIL"),     GetResource("L_ParamsRemoveInstalledIL"),     GetResource("L_optRemoveInstalledILUsage")
    OptLine GetResource("L_optClearTkaOnly"),          "",                                           GetResource("L_optClearTkaOnlyUsage")
    OptLine GetResource("L_optSetTkaOnly"),            "",                                           GetResource("L_optSetTkaOnlyUsage")
    OptLine GetResource("L_optListTkaCerts"),          "",                                           GetResource("L_optListTkaCertsUsage")
    OptLine GetResource("L_optForceTkaActivation"),    GetResource("L_ParamsForceTkaActivation"),    GetResource("L_optForceTkaActivationUsage")

    LineFlush ""

    LineOut GetResource("L_MsgKmsOptions")
    OptLine GetResource("L_optSetKmsListenPort"),      GetResource("L_ParamsSetListenKmsPort"),      GetResource("L_optSetKmsListenPortUsage")
    OptLine GetResource("L_optSetActivationInterval"), GetResource("L_ParamsSetActivationInterval"), GetResource("L_optSetActivationIntervalUsage")
    OptLine GetResource("L_optSetRenewalInterval"),    GetResource("L_ParamsSetRenewalInterval"),    GetResource("L_optSetRenewalIntervalUsage")
    OptLine GetResource("L_optSetDNS"),                "",                                           GetResource("L_optSetDNSUsage")
    OptLine GetResource("L_optClearDNS"),              "",                                           GetResource("L_optClearDNSUsage")
    OptLine GetResource("L_optSetNormalPriority"),     "",                                           GetResource("L_optSetNormalPriorityUsage")
    OptLine GetResource("L_optClearNormalPriority"),   "",                                           GetResource("L_optClearNormalPriorityUsage")

    ExitScript 1
End Sub

Private Sub OptLine(strOption, strParams, strUsage)
    LineOut "/" & strOption & " " & strParams
    LineOut "    " & strUsage
End Sub

Private Sub OptLine2(strOption, strParam1, strParam2, strUsage)
    LineOut "/" & strOption & " " & strParam1 & " " & strParam2
    LineOut "    " & strUsage
End Sub

Private Sub ExecCommandLine
    Dim intOption, indexOption
    Dim strOption, chOpt
    Dim remoteInfo(3)

    '
    ' First three parameters before "/" or "-" may be remote connection info
    '

    remoteInfo(0) = "."
    intOption = intUnknownOption

    For indexOption = 0 To 3
        If indexOption >= WScript.Arguments.Count Then
            Exit For
        End If

        strOption = WScript.Arguments.Item(indexOption)

        chOpt = Left(strOption, 1)
        If chOpt = "/" Or chOpt = "-" Then
            intOption = intKnownOption
            Exit For
        End If

        remoteInfo(indexOption) = strOption
    Next

    '
    ' Connect to remote only if syntax is reasonably good
    '

    If intUnknownOption = intOption Or 2 = indexOption Then
        g_strComputer = "."
        intOption = intUnknownOption
    Else
        g_strComputer = remoteInfo(0)
        g_strUserName = remoteInfo(1)
        g_strPassword = remoteInfo(2)
    End If

    If Not booleanConnect() Then
        ExitScript 1
    End If

    If intUnknownOption = intOption Then
        LineOut GetResource("L_MsgInvalidOptions")
        LineOut ""
        Call DisplayUsage()
    End If

    intOption = ParseCommandLine(indexOption)

    If intUnknownOption = intOption Then
        LineOut GetResource("L_MsgUnrecognizedOption") & WScript.Arguments.Item(indexOption)
        LineOut ""
        Call DisplayUsage()
    End If
End Sub

Private Function ParseCommandLine(index)
    Dim strOption, chOpt

    ParseCommandLine = intKnownOption

    strOption = LCase(WScript.Arguments.Item(index))

    chOpt = Left(strOption, 1)

    If (chOpt <> "-") And (chOpt <> "/") Then
        ParseCommandLine = intUnknownOption
        Exit Function
    End If

    strOption = Right(strOption, Len(strOption) - 1)

    If strOption = GetResource("L_optInstallLicense") Then

        If HandleOptionParam(index+1, True, GetResource("L_optInstallLicense"), GetResource("L_ParamsLicenseFile")) Then
            InstallLicense WScript.Arguments.Item(index+1)
        End If

    ElseIf strOption = GetResource("L_optInstallProductKey") Then

        If HandleOptionParam(index+1, True, GetResource("L_optInstallProductKey"), GetResource("L_ParamsProductKey")) Then
            InstallProductKey WScript.Arguments.Item(index+1)
        End If

    ElseIf strOption = GetResource("L_optUninstallProductKey") Then

        If HandleOptionParam(index+1, False, GetResource("L_optUninstallProductKey"), GetResource("L_ParamsActivationIDOptional")) Then
            UninstallProductKey WScript.Arguments.Item(index+1)
        Else
            UninstallProductKey ""
        End If

    ElseIf strOption = GetResource("L_optDisplayIID") Then

        If HandleOptionParam(index+1, False, GetResource("L_optDisplayIID"), GetResource("L_ParamsActivationIDOptional")) Then
            DisplayIID WScript.Arguments.Item(index+1)
        Else
            DisplayIID ""
        End If

    ElseIf strOption = GetResource("L_optActivateProduct") Then

        If HandleOptionParam(index+1, False, GetResource("L_optActivateProduct"), GetResource("L_ParamsActivationIDOptional")) Then
            ActivateProduct WScript.Arguments.Item(index+1)
        Else
            ActivateProduct ""
        End If

    ElseIf strOption = GetResource("L_optPhoneActivateProduct") Then

        If HandleOptionParam(index+1, True, GetResource("L_optPhoneActivateProduct"), GetResource("L_ParamsPhoneActivate")) Then
            If HandleOptionParam(index+2, False, GetResource("L_optPhoneActivateProduct"), GetResource("L_ParamsActivationIDOptional")) Then
                PhoneActivateProduct WScript.Arguments.Item(index+1), WScript.Arguments.Item(index+2)
            Else
                PhoneActivateProduct WScript.Arguments.Item(index+1), ""
            End If
        End If

    ElseIf strOption = GetResource("L_optDisplayInformation") Then

        If HandleOptionParam(index+1, False, GetResource("L_optDisplayInformation"), "") Then
            DisplayAllInformation WScript.Arguments.Item(index+1), False
        Else
            DisplayAllInformation "", False
        End If

    ElseIf strOption = GetResource("L_optDisplayInformationVerbose") Then

        If HandleOptionParam(index+1, False, GetResource("L_optDisplayInformationVerbose"), "") Then
            DisplayAllInformation WScript.Arguments.Item(index+1), True
        Else
            DisplayAllInformation "", True
        End If

    ElseIf strOption = GetResource("L_optClearPKeyFromRegistry") Then

        ClearPKeyFromRegistry

    ElseIf strOption = GetResource("L_optReinstallLicenses") Then

        ReinstallLicenses

    ElseIf strOption = GetResource("L_optReArmWindows") Then

        ReArmWindows()

    ElseIf strOption = GetResource("L_optExpirationDatime") Then

        If HandleOptionParam(index+1, False, GetResource("L_optExpirationDatime"), GetResource("L_ParamsActivationIDOptional")) Then
            ExpirationDatime WScript.Arguments.Item(index+1)
        Else
            ExpirationDatime ""
        End If

    ElseIf strOption = GetResource("L_optSetKmsName") Then

        If HandleOptionParam(index+1, True, GetResource("L_optSetKmsName"), GetResource("L_ParamsSetKms")) Then
            If HandleOptionParam(index+2, False, GetResource("L_optSetKmsName"), GetResource("L_ParamsActivationIDOptional")) Then
                SetKmsMachineName WScript.Arguments.Item(index+1), WScript.Arguments.Item(index+2)
            Else
                SetKmsMachineName WScript.Arguments.Item(index+1), ""
            End If
        End If

    ElseIf strOption = GetResource("L_optClearKmsName") Then

        If HandleOptionParam(index+1, False, GetResource("L_optClearKmsName"), GetResource("L_ParamsActivationIDOptional")) Then
            ClearKms WScript.Arguments.Item(index+1)
        Else
            ClearKms ""
        End If

    ElseIf strOption = GetResource("L_optSetKmsHostCaching") Then

        SetHostCachingDisable(False)        

    ElseIf strOption = GetResource("L_optClearKmsHostCaching") Then

        SetHostCachingDisable(True)        

    ElseIf strOption = GetResource("L_optSetActivationInterval") Then

        If HandleOptionParam(index+1, True, GetResource("L_optSetActivationInterval"), GetResource("L_ParamsSetActivationInterval")) Then
            SetActivationInterval  WScript.Arguments.Item(index+1)
        End If

    ElseIf strOption = GetResource("L_optSetRenewalInterval") Then

        If HandleOptionParam(index+1, True, GetResource("L_optSetRenewalInterval"), GetResource("L_ParamsSetRenewalInterval")) Then
            SetRenewalInterval  WScript.Arguments.Item(index+1)
        End If

    ElseIf strOption = GetResource("L_optSetKmsListenPort") Then

        If HandleOptionParam(index+1, True, GetResource("L_optSetKmsListenPort"), GetResource("L_ParamsSetListenKmsPort")) Then
            SetKmsListenPort WScript.Arguments.Item(index+1)
        End If

    ElseIf strOption = GetResource("L_optSetDNS") Then

        SetDnsPublishingDisabled(False)

    ElseIf strOption = GetResource("L_optClearDNS") Then

        SetDnsPublishingDisabled(True)

    ElseIf strOption = GetResource("L_optSetNormalPriority") Then

        SetKmsLowPriority(False)

    ElseIf strOption = GetResource("L_optClearNormalPriority") Then

        SetKmsLowPriority(True)

    ElseIf strOption = GetResource("L_optListInstalledILs") Then

        TkaListILs

    ElseIf strOption = GetResource("L_optRemoveInstalledIL") Then

        If HandleOptionParam(index+2, True, GetResource("L_optRemoveInstalledIL"), GetResource("L_ParamsRemoveInstalledIL")) Then
            TkaRemoveIL WScript.Arguments.Item(index+1), WScript.Arguments.Item(index+2)
        End If

    ElseIf strOption = GetResource("L_optClearTkaOnly") Then

        TkaSetTokenActivationOnly False

    ElseIf strOption = GetResource("L_optSetTkaOnly") Then

        TkaSetTokenActivationOnly True

    ElseIf strOption = GetResource("L_optListTkaCerts") Then

        TkaListCerts

    ElseIf strOption = GetResource("L_optForceTkaActivation") Then

        If HandleOptionParam(index+2, False, GetResource("L_optForceTkaActivation"), GetResource("L_ParamsForceTkaActivation")) Then
            TkaActivate WScript.Arguments.Item(index+1), WScript.Arguments.Item(index+2)
        ElseIf HandleOptionParam(index+1, True, GetResource("L_optForceTkaActivation"), GetResource("L_ParamsForceTkaActivation")) Then
            TkaActivate WScript.Arguments.Item(index+1), ""
        End If

    Else

        ParseCommandLine = intUnknownOption

    End If

End Function

' global options

Private Function CheckProductForCommand(objProduct, strActivationID)
    Dim bCheckProductForCommand

    bCheckProductForCommand = False

    If (strActivationID = "" And LCase(objProduct.ApplicationId) = WindowsAppId And (objProduct.LicenseIsAddon = False)) Then
        bCheckProductForCommand = True
    End If

    If (LCase(objProduct.ID) = strActivationID) Then
        bCheckProductForCommand = True
    End If
    
    CheckProductForCommand = bCheckProductForCommand
End Function

Private Sub UninstallProductKey(strActivationID)
    Dim objService, objProduct
    Dim lRet, strVersion, strDescription
    Dim kmsServerFound, uninstallDone
    Dim iIsPrimaryWindowsSku, bPrimaryWindowsSkuKeyUninstalled
    Dim bCheckProductForCommand

    On Error Resume Next

    strActivationID = LCase(strActivationID)
    kmsServerFound = False
    uninstallDone = False

    set objService = GetServiceObject("Version")
    strVersion = objService.Version

    For Each objProduct in GetProductCollection(ProductIsPrimarySkuSelectClause & ", ProductKeyID", PartialProductKeyNonNullWhereClause)
        strDescription = objProduct.Description

        bCheckProductForCommand = CheckProductForCommand(objProduct, strActivationID)

        If (bCheckProductForCommand) Then
            iIsPrimaryWindowsSku = GetIsPrimaryWindowsSKU(objProduct)
            If (strActivationID = "") And (iIsPrimaryWindowsSku = 2) Then
                    OutputIndeterminateOperationWarning(objProduct)
            End If

            objProduct.UninstallProductKey()
            QuitIfError()
            
            ' Uninstalling a product key could change Windows licensing state. 
            ' Since the service determines if it can shut down and when is the next start time
            ' based on the licensing state we should reconsume the licenses here.    
            objService.RefreshLicenseStatus()

            ' For Windows (i.e. if no activationID specified), always
            ' ensure that product-key for primary SKU is uninstalled
            If (strActivationID <> "") Or (iIsPrimaryWindowsSku = 1) Then
                uninstallDone = True
            End If

            LineOut GetResource("L_MsgUninstalledPKey")

        ' Check whether a ActID belongs to KMS server.
        ' Do this for all ActID other than one whose pkey is being uninstalled
        ElseIf IsKmsServer(strDescription) Then
            kmsServerFound = True
        End If

        If (kmsServerFound = True) And (uninstallDone = True) Then
            Exit For
        End If
    Next

    If kmsServerFound = True Then
        ' Set the KMS version in the registry (both 64 and 32 bit locations)
        lRet = SetRegistryStr(HKEY_LOCAL_MACHINE, SLKeyPath, "KeyManagementServiceVersion", strVersion)
        If (lRet <> 0) Then
            QuitWithError Hex(lRet)
        End If

        lRet = SetRegistryStr(HKEY_LOCAL_MACHINE, SLKeyPath32, "KeyManagementServiceVersion", strVersion)
        If (lRet <> 0) Then
            QuitWithError Hex(lRet)
        End If
    Else
        ' Clear the KMS version from the registry (both 64 and 32 bit locations)
        lRet = DeleteRegistryValue(HKEY_LOCAL_MACHINE, SLKeyPath, "KeyManagementServiceVersion")
        If (lRet <> 0 And lRet <> 2) Then
            QuitWithError Hex(lRet)
        End If

        lRet = DeleteRegistryValue(HKEY_LOCAL_MACHINE, SLKeyPath32, "KeyManagementServiceVersion")
        If (lRet <> 0 And lRet <> 2) Then
            QuitWithError Hex(lRet)
        End If
    End If

    If uninstallDone = False Then
        LineOut GetResource("L_MsgErrorPKey")
    End If
End Sub

Private Sub DisplayIID(strActivationID)
    Dim objProduct
    Dim iIsPrimaryWindowsSku, bFoundAtLeastOneKey
    Dim bCheckProductForCommand

    strActivationID = LCase(strActivationID)

    bFoundAtLeastOneKey = False
    For Each objProduct in GetProductCollection(ProductIsPrimarySkuSelectClause & ", OfflineInstallationId", PartialProductKeyNonNullWhereClause)

        bCheckProductForCommand = CheckProductForCommand(objProduct, strActivationID)

        If (bCheckProductForCommand) Then
            iIsPrimaryWindowsSku = GetIsPrimaryWindowsSKU(objProduct)
            If (strActivationID = "") And (iIsPrimaryWindowsSku = 2) Then
                    OutputIndeterminateOperationWarning(objProduct)
            End If

            LineOut GetResource("L_MsgInstallationID") & objProduct.OfflineInstallationId
            bFoundAtLeastOneKey = True

            If (strActivationID <> "") Or (iIsPrimaryWindowsSku = 1) Then
                Exit Sub
            End If
        End If
    Next

    If (bFoundAtLeastOneKey = TRUE) Then
        LineOut ""
        LineOut GetResource("L_MsgPhoneNumbers")
    Else
        LineOut GetResource("L_MsgErrorProductNotFound")
    End If
End Sub

Private Sub DisplayActivatingSku(objProduct)
    Dim strOutput
    
    strOutput = Replace(GetResource("L_MsgActivating"), "%PRODUCTNAME%", objProduct.Name)
    strOutput = Replace(strOutput, "%PRODUCTID%", objProduct.ID)
    LineFlush strOutput
End Sub

Private Sub DisplayActivatedStatus(objProduct)
    If (objProduct.LicenseStatus = 1) Then
        LineOut GetResource("L_MsgActivated")
    ElseIf (objProduct.LicenseStatus = 4) Then
        LineOut GetResource("L_MsgErrorText_8") & GetResource("L_MsgErrorText_11")
    ElseIf ((objProduct.LicenseStatus = 5) And (objProduct.LicenseStatusReason = HR_SL_E_NOT_GENUINE)) Then
        LineOut GetResource("L_MsgErrorText_8") & GetResource("L_MsgErrorText_12")
    ElseIf (objProduct.LicenseStatus = 6) Then
        LineOut GetResource("L_MsgActivated")
        LineOut GetResource("L_MsgLicenseStatusExtendedGrace_1")
    Else
        LineOut GetResource("L_MsgActivated_Failed")
    End If
End Sub

Private Sub ActivateProduct(strActivationID)
    Dim objService, objProduct
    Dim iIsPrimaryWindowsSku, bFoundAtLeastOneKey
    Dim strOutput
    Dim bCheckProductForCommand

    strActivationID = LCase(strActivationID)

    bFoundAtLeastOneKey = False

    set objService = GetServiceObject("Version")

    For Each objProduct in GetProductCollection(ProductIsPrimarySkuSelectClause & ", LicenseStatus", PartialProductKeyNonNullWhereClause)

        bCheckProductForCommand = CheckProductForCommand(objProduct, strActivationID)

        If (bCheckProductForCommand) Then
            iIsPrimaryWindowsSku = GetIsPrimaryWindowsSKU(objProduct)
            If (strActivationID = "") And (iIsPrimaryWindowsSku = 2) Then
                    OutputIndeterminateOperationWarning(objProduct)
            End If

            strOutput = Replace(GetResource("L_MsgActivating"), "%PRODUCTNAME%", objProduct.Name)
            strOutput = Replace(strOutput, "%PRODUCTID%", objProduct.ID)
            LineOut strOutput
            On Error Resume Next
            objProduct.Activate()
            QuitIfError()
            objService.RefreshLicenseStatus()
            objProduct.refresh_
            DisplayActivatedStatus objProduct

            bFoundAtLeastOneKey = True
            If (strActivationID <> "") Or (iIsPrimaryWindowsSku = 1) Then
                Exit Sub
            End If
        End If
    Next

    If (bFoundAtLeastOneKey = True) Then
        Exit Sub
    End If

    LineOut GetResource("L_MsgErrorProductNotFound")
End Sub

Private Sub PhoneActivateProduct(strCID, strActivationID)
    Dim objService, objProduct
    Dim iIsPrimaryWindowsSku, bFoundAtLeastOneKey
    Dim strOutput
    Dim bCheckProductForCommand

    strActivationID = LCase(strActivationID)

    bFoundAtLeastOneKey = False
    set objService = GetServiceObject("Version")

    For Each objProduct in GetProductCollection(ProductIsPrimarySkuSelectClause & ", OfflineInstallationId, LicenseStatus, LicenseStatusReason", PartialProductKeyNonNullWhereClause)

        bCheckProductForCommand = CheckProductForCommand(objProduct, strActivationID)

        If (bCheckProductForCommand) Then
            iIsPrimaryWindowsSku = GetIsPrimaryWindowsSKU(objProduct)
            If (strActivationID = "") And (iIsPrimaryWindowsSku = 2) Then
                    OutputIndeterminateOperationWarning(objProduct)
            End If

            On Error Resume Next
            objProduct.DepositOfflineConfirmationId objProduct.OfflineInstallationId, strCID
            QuitIfError()
            objService.RefreshLicenseStatus()
            objProduct.refresh_
            If (objProduct.LicenseStatus = 1) Then
                strOutput = Replace(GetResource("L_MsgConfID"), "%ACTID%", objProduct.ID)
                LineOut strOutput
            ElseIf (objProduct.LicenseStatus = 4) Then
                LineOut GetResource("L_MsgErrorText_8") & GetResource("L_MsgErrorText_11")
            ElseIf ((objProduct.LicenseStatus = 5) And (objProduct.LicenseStatusReason = HR_SL_E_NOT_GENUINE)) Then
                    LineOut GetResource("L_MsgErrorText_8") & GetResource("L_MsgErrorText_12")                    
            ElseIf (objProduct.LicenseStatus = 6) Then
                    LineOut GetResource("L_MsgActivated")
                    LineOut GetResource("L_MsgLicenseStatusExtendedGrace_1")
            Else
                LineOut GetResource("L_MsgActivated_Failed")
            End If

            bFoundAtLeastOneKey = True
            If (strActivationID <> "") Or (iIsPrimaryWindowsSku = 1) Then
                Exit Sub
            End If                    
        End If
    Next

    If (bFoundAtLeastOneKey = True) Then
        Exit Sub
    End If

    LineOut GetResource("L_MsgErrorProductNotFound")
End Sub

Private Sub DisplayKMSInformation(objService, objProduct)
    Dim dwValue
    Dim boolValue
    Dim KeyManagementServiceTotalRequests

    Dim objProductKMSValues

    set objProductKMSValues = GetProductObject( _
        "IsKeyManagementServiceMachine, KeyManagementServiceCurrentCount, " & _
        "KeyManagementServiceTotalRequests, KeyManagementServiceFailedRequests, " & _
        "KeyManagementServiceUnlicensedRequests, KeyManagementServiceLicensedRequests, " & _
        "KeyManagementServiceOOBGraceRequests, KeyManagementServiceOOTGraceRequests, " & _
        "KeyManagementServiceNonGenuineGraceRequests, KeyManagementServiceNotificationRequests", _
        "id = '" & objProduct.ID & "'")

    If objProductKMSValues.IsKeyManagementServiceMachine > 0 Then
        LineOut ""
        LineOut GetResource("L_MsgKmsEnabled")
        LineOut "    " & GetResource("L_MsgKmsCurrentCount") & objProductKMSValues.KeyManagementServiceCurrentCount

        dwValue = objService.KeyManagementServiceListeningPort 
        If 0 = dwValue Then
            LineOut "    " & GetResource("L_MsgKmsListeningOnPort") & DefaultPort
        Else
            LineOut "    " & GetResource("L_MsgKmsListeningOnPort") & dwValue
        End If

        boolValue = objService.KeyManagementServiceDnsPublishing 
        If true = boolValue Then
            LineOut "    " & GetResource("L_MsgKmsDnsPublishingEnabled")
        Else
            LineOut "    " & GetResource("L_MsgKmsDnsPublishingDisabled")
        End If

        boolValue = objService.KeyManagementServiceLowPriority 
        If false = boolValue Then
            LineOut "    " & GetResource("L_MsgKmsPriNormal")
        Else
            LineOut "    " & GetResource("L_MsgKmsPriLow")
        End If

        On Error Resume Next

        KeyManagementServiceTotalRequests = objProductKMSValues.KeyManagementServiceTotalRequests

        If (Not(IsNull(KeyManagementServiceTotalRequests))) And (Not(IsEmpty(KeyManagementServiceTotalRequests))) Then
            LineOut ""
            LineOut GetResource("L_MsgKmsCumulativeRequestsFromClients")
            LineOut "    " & GetResource("L_MsgKmsTotalRequestsRecieved") & objProductKMSValues.KeyManagementServiceTotalRequests
            LineOut "    " & GetResource("L_MsgKmsFailedRequestsReceived") & objProductKMSValues.KeyManagementServiceFailedRequests
            LineOut "    " & GetResource("L_MsgKmsRequestsWithStatusUnlicensed") & objProductKMSValues.KeyManagementServiceUnlicensedRequests
            LineOut "    " & GetResource("L_MsgKmsRequestsWithStatusLicensed") & objProductKMSValues.KeyManagementServiceLicensedRequests
            LineOut "    " & GetResource("L_MsgKmsRequestsWithStatusInitialGrace") & objProductKMSValues.KeyManagementServiceOOBGraceRequests
            LineOut "    " & GetResource("L_MsgKmsRequestsWithStatusLicenseExpiredOrHwidOot") & objProductKMSValues.KeyManagementServiceOOTGraceRequests
            LineOut "    " & GetResource("L_MsgKmsRequestsWithStatusNonGenuineGrace") & objProductKMSValues.KeyManagementServiceNonGenuineGraceRequests
            LineOut "    " & GetResource("L_MsgKmsRequestsWithStatusNotification") & objProductKMSValues.KeyManagementServiceNotificationRequests            
        End If
    End If
End Sub

Private Sub DisplayTkaClientInformation(objService, objProduct)
    LineOut ""
    LineOut GetResource("L_MsgTkaInfo")

    LineOut "    " & Replace(GetResource("L_MsgTkaInfoILID"      ), "%ILID%"      , objProduct.TokenActivationILID)
    LineOut "    " & Replace(GetResource("L_MsgTkaInfoILVID"     ), "%ILVID%"     , objProduct.TokenActivationILVID)
    LineOut "    " & Replace(GetResource("L_MsgTkaInfoGrantNo"   ), "%GRANTNO%"   , objProduct.TokenActivationGrantNumber)
    LineOut "    " & Replace(GetResource("L_MsgTkaInfoThumbprint"), "%THUMBPRINT%", objProduct.TokenActivationCertificateThumbprint)
End Sub

Private Sub DisplayKMSClientInformation(objService, objProduct)
    Dim strKms, strPort, strOutput
    Dim iVLRenewalInterval, iVLActivationInterval

    iVLRenewalInterval = objProduct.VLRenewalInterval
    iVLActivationInterval = objProduct.VLActivationInterval

    LineOut ""
    LineOut GetResource("L_MsgKmsInfo")
    LineOut "    " & GetResource("L_MsgCmid") & objService.ClientMachineID

    strKms = objProduct.KeyManagementServiceMachine

    if strKms <> "" And Not IsNull(strKms) Then
        strPort = objProduct.KeyManagementServicePort
        If (strPort = 0) Then
            strPort = DefaultPort
        End If
        LineOut "    " & GetResource("L_MsgRegisteredKmsName") & strKms & ":" & strPort
    Else
        strKms = objProduct.DiscoveredKeyManagementServiceMachineName
        strPort = objProduct.DiscoveredKeyManagementServiceMachinePort
        
        If IsNull(strKms) Or (strKms = "") Or IsNull(strPort) Or (strPort = 0) Then
            LineOut "    " & GetResource("L_MsgKmsFromDnsUnavailable")
        Else
            LineOut "    " & GetResource("L_MsgKmsFromDns") & strKms & ":" & strPort
        End If
    End If

    LineOut "    " & GetResource("L_MsgKmsPID4") & objProduct.KeyManagementServiceProductKeyID
    strOutput = Replace(GetResource("L_MsgActivationInterval"), "%INTERVAL%", iVLActivationInterval)
    LineOut "    " & strOutput
    strOutput = Replace(GetResource("L_MsgRenewalInterval"), "%INTERVAL%", iVLRenewalInterval)
    LineOut "    " & strOutput

        if (objService.KeyManagementServiceHostCaching = True) Then
            LineOut "    " & GetResource("L_MsgKmsHostCachingEnabled")  
        Else
            LineOut "    " & GetResource("L_MsgKmsHostCachingDisabled") 
        End If
End Sub

'
' Display all information for /dlv and /dli
' If you add need to access new properties through WMI you must add them to the 
' queries for service/object.  Be sure to check that the object properties in DisplayAllInformation()
' are requested for function/methods such as GetIsPrimaryWindowsSKU() and DisplayKMSClientInformation().
'
Private Sub DisplayAllInformation(strParm, bVerbose)
    Dim objService, objProduct
    Dim strServiceSelectClause
    Dim strDescription, bKmsClient, strSLActID, bKmsServer, bTBL
    Dim ls, gpMin, gpDay, displayDate
    Dim strOutput
    Dim bShowSkuInformation
    Dim iIsPrimaryWindowsSku, bUseDefault
    Dim productKeyFound

    Dim strErr
    strParm = LCase(strParm)
    productKeyFound = False
    strServiceSelectClause = "KeyManagementServiceListeningPort, KeyManagementServiceDnsPublishing, " & _
                                "KeyManagementServiceLowPriority, ClientMachineId, KeyManagementServiceHostCaching, " & _
                                "Version"

    If bVerbose Then
        strServiceSelectClause = "RemainingWindowsReArmCount, " & strServiceSelectClause
    End If

    set objService = GetServiceObject(strServiceSelectClause)

    If bVerbose Then
        LineOut GetResource("L_MsgServiceVersion") & objService.Version
    End If

    For Each objProduct in GetProductCollection(ProductIsPrimarySkuSelectClause & ", " & _
                            "ProductKeyID, OfflineInstallationId, " & _ 
                            "ProcessorURL, MachineURL, UseLicenseURL, ProductKeyURL, " & _
                            "GracePeriodRemaining, LicenseStatus, LicenseStatusReason, EvaluationEndDate, " & _
                            "VLRenewalInterval, VLActivationInterval, KeyManagementServiceMachine, " & _
                            "KeyManagementServicePort, DiscoveredKeyManagementServiceMachineName, " & _
                            "DiscoveredKeyManagementServiceMachinePort, KeyManagementServiceProductKeyID," & _
                            "TokenActivationILID, TokenActivationILVID, TokenActivationGrantNumber," & _
                            "TokenActivationCertificateThumbprint, TokenActivationAdditionalInfo, TrustedTime", _
                            EmptyWhereClause)

        strSLActID = objProduct.ID

        ' Display information if:
        '    parm = "all" or
        '    ActID = parm or
        '    default to current ActID (parm = "" and IsPrimaryWindowsSKU is 1 or 2)
        iIsPrimaryWindowsSku = GetIsPrimaryWindowsSKU(objProduct)
        bUseDefault = False
        bShowSkuInformation = False

        If (strParm = "" And ((iIsPrimaryWindowsSku = 1) Or (iIsPrimaryWindowsSku = 2))) Then
            bUseDefault = True
            bShowSkuInformation = True
        End If

        If (strParm = "" And (objProduct.LicenseIsAddon And objProduct.PartialProductKey <> "")) Then
            bShowSkuInformation = True
        End If

        If (strParm = "all") Then
            bShowSkuInformation = True
        End If

        If (strParm = LCase(strSLActID)) Then
            bShowSkuInformation = True
        End If

        If (bShowSkuInformation) Then

            strDescription = objProduct.Description

            'If the user didn't specify anything and we are showing the default case, warn them
            ' if this can't be verified as the primary SKU
            If ((bUseDefault = True) And (iIsPrimaryWindowsSku = 2)) Then
                OutputIndeterminateOperationWarning(objProduct)
            End IF

            productKeyFound = True

            LineOut ""
            LineOut GetResource("L_MsgProductName") & objProduct.Name

            LineOut GetResource("L_MsgProductDesc") & strDescription

            If objProduct.TokenActivationAdditionalInfo <> "" Then
                LineOut Replace( _
                    GetResource("L_MsgTkaInfoAdditionalInfo"), _
                    "%MOREINFO%", _
                    objProduct.TokenActivationAdditionalInfo _
                    )
            End If

            bKmsServer = IsKmsServer(strDescription)
            bKmsClient = IsKmsClient(strDescription)
            bTBL       = IsTBL(strDescription)

            If bVerbose Then
                LineOut GetResource("L_MsgActID") & strSLActID
                LineOut GetResource("L_MsgAppID") & objProduct.ApplicationID
                LineOut GetResource("L_MsgPID4") & objProduct.ProductKeyID
                LineOut GetResource("L_MsgInstallationID") & objProduct.OfflineInstallationId

                If NOT bKmsClient Then
                    LineOut GetResource("L_MsgProcessorCertUrl") & objProduct.ProcessorURL
                    LineOut GetResource("L_MsgMachineCertUrl") & objProduct.MachineURL
                    LineOut GetResource("L_MsgUseLicenseCertUrl") & objProduct.UseLicenseURL
                    LineOut GetResource("L_MsgPKeyCertUrl") & objProduct.ProductKeyURL
                End If
            End If

            If objProduct.PartialProductKey <> "" Then
                LineOut GetResource("L_MsgPartialPKey") & objProduct.PartialProductKey
            Else
                LineOut GetResource("L_MsgErrorLicenseNotInUse")
            End If

            ls = objProduct.LicenseStatus

            If ls = 0 Then
                LineOut GetResource("L_MsgLicenseStatusUnlicensed_1")

            ElseIf ls = 1 Then
                LineOut GetResource("L_MsgLicenseStatusLicensed_1")
                gpMin = objProduct.GracePeriodRemaining
                If (gpMin <> 0) Then
                    gpDay = Int(gpMin / (24 * 60))
                    If (bTBL) Then
                        strOutput = Replace(GetResource("L_MsgLicenseStatusTBL_1"), "%MINUTE%", gpMin)
                    Else
                        strOutput = Replace(GetResource("L_MsgLicenseStatusVL_1"), "%MINUTE%", gpMin)
                    End If
                    strOutput = Replace(strOutput, "%DAY%", gpDay)
                    LineOut strOutput
                End If

            ElseIf ls = 2 Then
                LineOut GetResource("L_MsgLicenseStatusInitialGrace_1")
                gpMin = objProduct.GracePeriodRemaining
                gpDay = Int(gpMin / (24 * 60))
                strOutput = Replace(GetResource("L_MsgLicenseStatusTimeRemaining"), "%MINUTE%", gpMin)
                strOutput = Replace(strOutput, "%DAY%", gpDay)
                LineOut strOutput

            ElseIf ls = 3 Then
                LineOut GetResource("L_MsgLicenseStatusAdditionalGrace_1")
                gpMin = objProduct.GracePeriodRemaining
                gpDay = Int(gpMin / (24 * 60))
                strOutput = Replace(GetResource("L_MsgLicenseStatusTimeRemaining"), "%MINUTE%", gpMin)
                strOutput = Replace(strOutput, "%DAY%", gpDay)
                LineOut strOutput

            ElseIf ls = 4 Then
                LineOut GetResource("L_MsgLicenseStatusNonGenuineGrace_1")
                gpMin = objProduct.GracePeriodRemaining
                gpDay = Int(gpMin / (24 * 60))
                strOutput = Replace(GetResource("L_MsgLicenseStatusTimeRemaining"), "%MINUTE%", gpMin)
                strOutput = Replace(strOutput, "%DAY%", gpDay)
                LineOut strOutput
                
            ElseIf ls = 5 Then
                LineOut GetResource("L_MsgLicenseStatusNotification_1")
                strErr = CStr(Hex(objProduct.LicenseStatusReason))           
                if (objProduct.LicenseStatusReason = HR_SL_E_NOT_GENUINE) Then
                   strOutput = Replace(GetResource("L_MsgNotificationErrorReasonNonGenuine"), "%ERRCODE%", strErr)
                ElseIf (objProduct.LicenseStatusReason = HR_SL_E_GRACE_TIME_EXPIRED) Then
                    strOutput = Replace(GetResource("L_MsgNotificationErrorReasonExpiration"), "%ERRCODE%", strErr)
                Else
                    strOutput = Replace(GetResource("L_MsgNotificationErrorReasonOther"), "%ERRCODE%", strErr)                        
                End If
                LineOut strOutput

            ElseIf ls = 6 Then
                LineOut GetResource("L_MsgLicenseStatusExtendedGrace_1")
                gpMin = objProduct.GracePeriodRemaining
                gpDay = Int(gpMin / (24 * 60))
                strOutput = Replace(GetResource("L_MsgLicenseStatusTimeRemaining"), "%MINUTE%", gpMin)
                strOutput = Replace(strOutput, "%DAY%", gpDay)
                LineOut strOutput                    

            Else
                LineOut GetResource("L_MsgLicenseStatusUnknown")
            End If

            If (ls <> 0 And bVerbose) Then
                Set displayDate = CreateObject("WBemScripting.SWbemDateTime")
                displayDate.Value = objProduct.EvaluationEndDate
                If (displayDate.GetFileTime(false) <> 0) Then
                    LineOut GetResource("L_MsgLicenseStatusEvalEndData") & displayDate.GetVarDate
                End If
            End If

            If (bVerbose) Then

                LineOut Replace(GetResource("L_MsgRemainingWindowsRearmCount"), "%COUNT%", objService.RemainingWindowsReArmCount)

                Set displayDate = CreateObject("WBemScripting.SWbemDateTime")
                displayDate.Value = objProduct.TrustedTime
                If (displayDate.GetFileTime(false) <> 0) Then
                    LineOut GetResource("L_MsgCurrentTrustedTime") & displayDate.GetVarDate
                End If

            End If

            '
            ' KMS client properties
            '

            If bKmsClient Then
                If IsTokenActivated(objProduct) Then
                    DisplayTkaClientInformation objService, objProduct
                ElseIf ls <> 1 Then
                    LineOut GetResource("L_MsgPleaseActivateRefreshKMSInfo")
                Else
                    DisplayKMSClientInformation objService, objProduct
                End If
            End If

            If (bKmsServer Or (iIsPrimaryWindowsSku = 1) Or (iIsPrimaryWindowsSku = 2)) Then
                DisplayKMSInformation objService, objProduct
            End If

            'We should stop processing if we aren't processing All and either we were told to process a single
            'entry only or we found the primary SKU
            If strParm <> "all" Then
                If (strParm = LCase(strSLActID)) Then
                    Exit For  'no need to continue
                End If
            End If

            LineOut ""
        End If
    Next

    If productKeyFound = False Then
        LineOut GetResource("L_MsgErrorPKey")
    End If

End Sub

Private Sub InstallProductKey(strProductKey)
    Dim objService, objProduct
    Dim lRet, strDescription, strOutput, strVersion
    Dim iIsPrimaryWindowsSku, bIsKMS

    bIsKMS = False

    On Error Resume Next

    set objService = GetServiceObject("Version")
    strVersion = objService.Version
    objService.InstallProductKey(strProductKey)
    QuitIfError()
        
    ' Installing a product key could change Windows licensing state. 
    ' Since the service determines if it can shut down and when is the next start time
    ' based on the licensing state we should reconsume the licenses here.    
    objService.RefreshLicenseStatus()

    For Each objProduct in GetProductCollection(ProductIsPrimarySkuSelectClause, PartialProductKeyNonNullWhereClause)
        strDescription = objProduct.Description

        iIsPrimaryWindowsSku = GetIsPrimaryWindowsSKU(objProduct)
        If (iIsPrimaryWindowsSku = 2) Then
            OutputIndeterminateOperationWarning(objProduct)
        End If

        If IsKmsServer(strDescription) Then
            bIsKMS = True
            Exit For
        End If
    Next

    If (bIsKMS = True) Then
        ' Set the KMS version in the registry (64 and 32 bit versions)
        lRet = SetRegistryStr(HKEY_LOCAL_MACHINE, SLKeyPath, "KeyManagementServiceVersion", strVersion)
        If (lRet <> 0) Then
            QuitWithError Hex(lRet)
        End If

        If ExistsRegistryKey(HKEY_LOCAL_MACHINE, SLKeyPath32) Then
            lRet = SetRegistryStr(HKEY_LOCAL_MACHINE, SLKeyPath32, "KeyManagementServiceVersion", strVersion)
            If (lRet <> 0) Then
                QuitWithError Hex(lRet)
            End If
        End If
    Else
        ' Clear the KMS version in the registry (64 and 32 bit versions)
        lRet = DeleteRegistryValue(HKEY_LOCAL_MACHINE, SLKeyPath, "KeyManagementServiceVersion")
        If (lRet <> 0 And lRet <> 2 And lRet <> 5) Then
            QuitWithError Hex(lRet)
        End If

        lRet = DeleteRegistryValue(HKEY_LOCAL_MACHINE, SLKeyPath32, "KeyManagementServiceVersion")
        If (lRet <> 0 And lRet <> 2 And lRet <> 5) Then
            QuitWithError Hex(lRet)
        End If
    End If

    strOutput = Replace(GetResource("L_MsgInstalledPKey"), "%PKEY%", strProductKey)
    LineOut strOutput
End Sub

Private Sub OutputIndeterminateOperationWarning(objProduct)
    Dim strOutput

    LineOut GetResource("L_MsgUndeterminedPrimaryKeyOperation")
    strOutput = Replace(GetResource("L_MsgUndeterminedOperationFormat"), "%PRODUCTDESCRIPTION%", objProduct.Description)
    strOutput = Replace(strOutput, "%PRODUCTID%", objProduct.ID)
    LineOut strOutput
End Sub

Private Sub ClearPKeyFromRegistry()
    Dim objService

    On Error Resume Next

    set objService = GetServiceObject("Version")
    QuitIfError()
    
    objService.ClearProductKeyFromRegistry()
    QuitIfError()

    LineOut GetResource("L_MsgClearedPKey")
End Sub

Private Sub InstallLicenseFiles (strParentDirectory, fso)
    Dim file, files, folder, subFolder

    Set folder = fso.GetFolder(strParentDirectory)
    Set files = folder.Files

    ' Install all license files in folder
    For Each file In files
        If Right(file.Name, 7) = ".xrm-ms" Then
            InstallLicense strParentDirectory & "\" & file.Name
        End If
    Next

    For Each subFolder in folder.SubFolders
        InstallLicenseFiles subFolder, fso
    Next
End Sub

Private Sub ReinstallLicenses()
    Dim shell, fso, strOemFolder
    Dim strSppTokensFolder, folder, subFolder
    Set shell = WScript.CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")

    strOemFolder = shell.ExpandEnvironmentStrings("%SystemRoot%") & "\system32\oem"
    strSppTokensFolder = shell.ExpandEnvironmentStrings("%SystemRoot%") & "\system32\spp\tokens"
    
    LineOut GetResource("L_MsgReinstallingLicenses")
    
    Set folder = fso.GetFolder(strSppTokensFolder)
    
    For Each subFolder in folder.SubFolders
        InstallLicenseFiles subFolder, fso
    Next
    
    If (fso.FolderExists(strOemFolder)) Then
        InstallLicenseFiles strOemFolder, fso
    End If

    LineOut GetResource("L_MsgLicensesReinstalled")
End Sub

Private Sub ReArmWindows
    Dim objService

    set objService = GetServiceObject("Version")
    On Error Resume Next

    objService.ReArmWindows()
    QuitIfError()

    LineOut GetResource("L_MsgRearm_1")
    LineOut GetResource("L_MsgRearm_2")
End Sub

Private Sub ExpirationDatime(strActivationID)
    Dim objProduct
    Dim strSLActID, ls, graceRemaining, strEnds
    Dim strOutput
    Dim strDescription, bTBL
    Dim iIsPrimaryWindowsSku
    Dim bFound

    strActivationID = LCase(strActivationID)

    bFound = False

    For Each objProduct in GetProductCollection(ProductIsPrimarySkuSelectClause & ", LicenseStatus, GracePeriodRemaining", EmptyWhereClause)
        If ((strActivationID = "" And LCase(objProduct.ApplicationId) = WindowsAppId) Or LCase(objProduct.ID) = strActivationID) And (objProduct.PartialProductKey <> "") Then
            strSLActID = objProduct.ID
            ls = objProduct.LicenseStatus
            graceRemaining = objProduct.GracePeriodRemaining
            strEnds = DateAdd("n", graceRemaining, Now)

            bFound = True

            iIsPrimaryWindowsSku = GetIsPrimaryWindowsSKU(objProduct)
            If (strActivationID = "") And (iIsPrimaryWindowsSku = 2) Then
                OutputIndeterminateOperationWarning(objProduct)
            End If

            strOutput = ""

            If ls = 0 Then
                strOutput = GetResource("L_MsgLicenseStatusUnlicensed")

            ElseIf ls = 1 Then
                If graceRemaining <> 0 Then           

                    strDescription = objProduct.Description

                    bTBL = IsTBL(strDescription)

                    If bTBL Then
                        strOutput = Replace(GetResource("L_MsgLicenseStatusTBL"), "%ENDDATE%", strEnds)
                    Else
                        strOutput = Replace(GetResource("L_MsgLicenseStatusVL"), "%ENDDATE%", strEnds)
                    End If
                Else
                	strOutput = GetResource("L_MsgLicenseStatusLicensed")
                End If

            ElseIf ls = 2 Then
                strOutput = Replace(GetResource("L_MsgLicenseStatusInitialGrace"), "%ENDDATE%", strEnds)
            ElseIf ls = 3 Then
                strOutput = Replace(GetResource("L_MsgLicenseStatusAdditionalGrace"), "%ENDDATE%", strEnds)
            ElseIf ls = 4 Then
                strOutput = Replace(GetResource("L_MsgLicenseStatusNonGenuineGrace"), "%ENDDATE%", strEnds)
            ElseIf ls = 5 Then
                strOutput =  GetResource("L_MsgLicenseStatusNotification")
            ElseIf ls = 6 Then
                strOutput = Replace(GetResource("L_MsgLicenseStatusExtendedGrace"), "%ENDDATE%", strEnds)
            End If

            If strOutput <> "" Then
                LineOut objProduct.Name & ":"
                Lineout "    " & strOutput
            End If
            
        End If
    Next

    If True <> bFound Then
        LineOut GetResource("L_MsgErrorPKey")
    End If
End Sub

' volume license service/client management


Private Sub QuitIfErrorRestoreKmsName(obj, strKmsName)
    Dim errNum

    If Err.Number <> 0 Then
        errNum  = Err.Number

        If strKmsName = "" Then
            obj.ClearKeyManagementServiceMachine()
        Else
            obj.SetKeyManagementServiceMachine(strKmsName)
        End If

        ShowErrorNum GetResource("L_MsgErrorText_8"), CStr(Hex(errNum))
        ExitScript errNum
    End If
End Sub

Private Sub SetKmsMachineName(strKmsNamePort, strActivationID)
    Dim objService, objProduct
    Dim nColon, strKmsName, strKmsNamePrev, strKmsPort, lRet, strOutput, nBracketEnd
    Dim nKmsPort
    Dim activationIDFound

    nBracketEnd = InStr(StrKmsNamePort, "]") 
    If InStr(strKmsNamePort, "[") = 1 And nBracketEnd > 1 Then
    ' IPV6 Address    
        If  Len(StrKmsNamePort) = nBracketEnd Then
            'No Port Number
            strKmsName = strKmsNamePort
            strKmsPort = ""
        Else
            strKmsName = Left(strKmsNamePort, nBracketEnd)
            strKmsPort = Right(strKmsNamePort, Len(strKmsNamePort) - nBracketEnd - 1)
        End If
    Else
    ' IPV4 Address
        nColon = InStr(1, strKmsNamePort, ":")
        If nColon <> 0 Then
            strKmsName = Left(strKmsNamePort, nColon - 1)
            strKmsPort = Right(strKmsNamePort, Len(strKmsNamePort) - nColon)
        Else
            strKmsName = strKmsNamePort
            strKmsPort = ""
        End If
    End If

    strActivationID = LCase(strActivationID)

    set objService = GetServiceObject("Version, KeyManagementServiceMachine")
    On Error Resume Next

    If (strActivationID = "") Then
        strKmsNamePrev = objService.KeyManagementServiceMachine

        If strKmsName <> "" Then
            objService.SetKeyManagementServiceMachine(strKmsName)
            QuitIfError()
        End If

        If strKmsPort <> "" Then
            nKmsPort = CLng(strKmsPort)
            QuitIfErrorRestoreKmsName objService, strKmsNamePrev
            objService.SetKeyManagementServicePort(nKmsPort)
            QuitIfErrorRestoreKmsName objService, strKmsNamePrev
        Else
            objService.ClearKeyManagementServicePort()
            QuitIfErrorRestoreKmsName objService, strKmsNamePrev
        End If
    Else
        activationIDFound = False

        For Each objProduct in GetProductCollection("ID, KeyManagementServiceMachine", EmptyWhereClause)
            If (LCase(objProduct.ID) = strActivationID) Then

                strKmsNamePrev = objProduct.KeyManagementServiceMachine

                If strKmsName <> "" Then
                    objProduct.SetKeyManagementServiceMachine(strKmsName)
                    QuitIfError()
                End If

                If strKmsPort <> "" Then
                    nKmsPort = CLng(strKmsPort)
                    QuitIfErrorRestoreKmsName objProduct, strKmsNamePrev
                    objProduct.SetKeyManagementServicePort(nKmsPort)
                    QuitIfErrorRestoreKmsName objProduct, strKmsNamePrev
                Else
                    objProduct.ClearKeyManagementServicePort()
                    QuitIfErrorRestoreKmsName objProduct, strKmsNamePrev
                End If

                activationIDFound = True
                Exit For
            End If    
        Next

        if activationIDFound = False Then
            strOutput = Replace(GetResource("L_MsgErrorActivationID"), "%ActID%", strActivationID)
            Lineout strOutput
        End If
    End If
    
    QuitIfError()

    If (strActivationID = "") Or (activationIDFound = True) Then
        strOutput = Replace(GetResource("L_MsgKmsNameSet"), "%KMS%", strKmsNamePort)
        LineOut strOutput
    End If
End Sub

Private Sub SetHostCachingDisable(boolHostCaching)
    Dim objService

    On Error Resume Next

    set objService = GetServiceObject("Version")
    QuitIfError()
    
    objService.DisableKeyManagementServiceHostCaching(boolHostCaching)
    QuitIfError()

    If boolHostCaching Then
        LineOut GetResource("L_MsgKmsHostCachingDisabled")
    Else
        LineOut GetResource("L_MsgKmsHostCachingEnabled")
    End If
    
End Sub

Private Sub ClearKms(strActivationID)
    Dim objService, objProduct
    Dim lRet
    Dim activationIDFound, strOutput

    strActivationID = LCase(strActivationID)

    On Error Resume Next

    set objService = GetServiceObject("Version")
    QuitIfError()

    If (strActivationID = "") Then
        objService.ClearKeyManagementServiceMachine()
        objService.ClearKeyManagementServicePort()
    Else
        activationIDFound = False

        For Each objProduct in GetProductCollection("ID", EmptyWhereClause)
            If (LCase(objProduct.ID) = strActivationID) Then
                objProduct.ClearKeyManagementServiceMachine()
                objProduct.ClearKeyManagementServicePort()

                activationIDFound = True
                Exit For
            End If
        Next

        if activationIDFound = False Then
            strOutput = Replace(GetResource("L_MsgErrorActivationID"), "%ActID%", strActivationID)
            Lineout strOutput
        End If
    End If

    If (strActivationID = "") Or (activationIDFound = True) Then
        LineOut GetResource("L_MsgKmsNameCleared")
    End If
End Sub

Private Sub SetActivationInterval(intInterval)
    Dim objService, objProduct
    Dim kmsFlag, strOutput

    If (intInterval < 0) Then
        LineOut GetResource("L_MsgInvalidDataError")
        Exit Sub
    End If

    On Error Resume Next

    set objService = GetServiceObject("Version")
    QuitIfError()

    For Each objProduct in GetProductCollection("ID, IsKeyManagementServiceMachine", PartialProductKeyNonNullWhereClause)
        kmsFlag = objProduct.IsKeyManagementServiceMachine
        If kmsFlag = 1 Then
            objService.SetVLActivationInterval(intInterval)
            QuitIfError()
            strOutput = Replace(GetResource("L_MsgActivationSet"), "%ACTIVATION%", intInterval)
            LineOut strOutput
            LineOut GetResource("L_MsgWarningKmsReboot")

            Exit For
        End If                
    Next

    If kmsFlag <> 1 Then
        LineOut GetResource("L_MsgWarningActivation")
    End If
End Sub

Private Sub SetRenewalInterval(intInterval)
    Dim objService, objProduct
    Dim kmsFlag, strOutput

    If (intInterval < 0) Then
        LineOut GetResource("L_MsgInvalidDataError")
        Exit Sub
    End If

    On Error Resume Next

    set objService = GetServiceObject("Version")
    QuitIfError()

    For Each objProduct in GetProductCollection("ID, IsKeyManagementServiceMachine", PartialProductKeyNonNullWhereClause)
        kmsFlag = objProduct.IsKeyManagementServiceMachine
        If kmsFlag Then
            objService.SetVLRenewalInterval(intInterval)
            QuitIfError()
            strOutput = Replace(GetResource("L_MsgRenewalSet"), "%RENEWAL%", intInterval)
            LineOut strOutput
            LineOut GetResource("L_MsgWarningKmsReboot")

            Exit For
        End If                
    Next

    If kmsFlag <> 1 Then
        LineOut GetResource("L_MsgWarningRenewal")
    End If
End Sub

Private Sub SetKmsListenPort(strPort)
    Dim objService, objProduct
    Dim kmsFlag, lRet, strOutput
    Dim nPort

    On Error Resume Next

    set objService = GetServiceObject("Version")
    QuitIfError()

    For Each objProduct in GetProductCollection("ID, IsKeyManagementServiceMachine", PartialProductKeyNonNullWhereClause)
        kmsFlag = objProduct.IsKeyManagementServiceMachine
        If kmsFlag Then
            nPort = CLng(strPort)
            objService.SetKeyManagementServiceListeningPort(nPort)
            QuitIfError()
            strOutput = Replace(GetResource("L_MsgKmsPortSet"), "%PORT%", strPort)
            LineOut strOutput
            LineOut GetResource("L_MsgWarningKmsReboot")

            Exit For
        End If
    Next

    If kmsFlag <> 1 Then
        LineOut GetResource("L_MsgWarningKmsPort")
    End If
End Sub

Private Sub SetDnsPublishingDisabled(bool)
    Dim objService, objProduct
    Dim kmsFlag, lRet, dwValue

    On Error Resume Next

    set objService = GetServiceObject("Version")
    QuitIfError()

    For Each objProduct in GetProductCollection("ID, IsKeyManagementServiceMachine", PartialProductKeyNonNullWhereClause)
        kmsFlag = objProduct.IsKeyManagementServiceMachine
        If kmsFlag Then
            objService.DisableKeyManagementServiceDnsPublishing(bool)
            QuitIfError()

            If bool Then
                LineOut GetResource("L_MsgKmsDnsPublishingDisabled")
            Else
                LineOut GetResource("L_MsgKmsDnsPublishingEnabled")
            End If
            LineOut GetResource("L_MsgWarningKmsReboot")

            Exit For
        End If
    Next

    If kmsFlag <> 1 Then
        LineOut GetResource("L_MsgKmsDnsPublishingWarning")
    End If
End Sub

Private Sub SetKmsLowPriority(bool)
    Dim objService, objProduct
    Dim kmsFlag, lRet, dwValue

    On Error Resume Next

    set objService = GetServiceObject("Version")
    QuitIfError()

    For Each objProduct in GetProductCollection("ID, IsKeyManagementServiceMachine", PartialProductKeyNonNullWhereClause)
        kmsFlag = objProduct.IsKeyManagementServiceMachine
        If kmsFlag Then
            objService.EnableKeyManagementServiceLowPriority(bool)
            QuitIfError()

            If bool Then
                LineOut GetResource("L_MsgKmsPriSetToLow")
            Else
                LineOut GetResource("L_MsgKmsPriSetToNormal")
            End If
            LineOut GetResource("L_MsgWarningKmsReboot")
        End If

        Exit For
    Next    


    If kmsFlag <> 1 Then
       LineOut GetResource("L_MsgWarningKmsPri")
    End If
End Sub

''
'' Token-based Activation Commands
''

Private Function IsTokenActivated(objProduct)

    Dim nILVID

    On Error Resume Next

    nILVID = objProduct.TokenActivationILVID

    IsTokenActivated = ((Err.Number = 0) And (nILVID <> &HFFFFFFFF))

End Function


Private Sub TkaListILs
    Dim objLicense
    Dim strHeader
    Dim strError
    Dim strGuids
    Dim arrGuids
    Dim nListed

    Dim objWmiDate

    LineOut GetResource("L_MsgTkaLicenses")
    LineOut ""

    Set objWmiDate = CreateObject("WBemScripting.SWbemDateTime")

    nListed = 0
    For Each objLicense in g_objWMIService.InstancesOf(TkaLicenseClass)

        strHeader = GetResource("L_MsgTkaLicenseHeader")
        strHeader = Replace(strHeader, "%ILID%" , objLicense.ILID )
        strHeader = Replace(strHeader, "%ILVID%", objLicense.ILVID)
        LineOut strHeader

        LineOut "    " & Replace(GetResource("L_MsgTkaLicenseILID"), "%ILID%", objLicense.ILID)
        LineOut "    " & Replace(GetResource("L_MsgTkaLicenseILVID"), "%ILVID%", objLicense.ILVID)

        If Not IsNull(objLicense.ExpirationDate) Then

            objWmiDate.Value = objLicense.ExpirationDate

            If (objWmiDate.GetFileTime(false) <> 0) Then
                LineOut "    " & Replace(GetResource("L_MsgTkaLicenseExpiration"), "%TODATE%", objWmiDate.GetVarDate)
            End If

        End If

        If Not IsNull(objLicense.AdditionalInfo) Then
            LineOut "    " & Replace(GetResource("L_MsgTkaLicenseAdditionalInfo"), "%MOREINFO%", objLicense.AdditionalInfo)
        End If

        If Not IsNull(objLicense.AuthorizationStatus) And _
           objLicense.AuthorizationStatus <> 0 _
        Then
            strError = CStr(Hex(objLicense.AuthorizationStatus))
            LineOut "    " & Replace(GetResource("L_MsgTkaLicenseAuthZStatus"), "%ERRCODE%", strError)
        Else            
            LineOut "    " & Replace(GetResource("L_MsgTkaLicenseDescr"), "%DESC%", objLicense.Description)
        End If
        
        LineOut ""
        nListed = nListed + 1
    Next

    if 0 = nListed Then
        LineOut GetResource("L_MsgTkaLicenseNone")
    End If
End Sub


Private Sub TkaRemoveIL(strILID, strILVID)
    Dim objLicense
    Dim strMsg
    Dim nRemoved

    Dim nILVID

    On Error Resume Next
    nILVID = CInt(strILVID)
    QuitIfError()

    LineOut GetResource("L_MsgTkaRemoving")
    LineOut ""

    nRemoved = 0
    For Each objLicense in g_objWMIService.InstancesOf(TkaLicenseClass)
        If strILID = objLicense.ILID And nILVID = objLicense.ILVID Then
            strMsg = GetResource("L_MsgTkaRemovedItem")
            strMsg = Replace(strMsg, "%SLID%", objLicense.ID)

            On Error Resume Next
            objLicense.Uninstall
            QuitIfError()
            LineOut strMsg
            nRemoved = nRemoved + 1
        End If
    Next

    If nRemoved = 0 Then
        LineOut GetResource("L_MsgTkaRemovedNone")
    End If
End Sub


Private Sub TkaSetTokenActivationOnly(bTao)
    Dim objService

    On Error Resume Next

    set objService = GetServiceObject("Version")
    QuitIfError()
    
    objService.DisableKeyManagementServiceActivation(bTao)
    QuitIfError()

    If bTao Then
        LineOut GetResource("L_MsgTkaTaoSet")
    Else
        LineOut GetResource("L_MsgTkaTaoClear")
    End If

End Sub

Private Sub TkaListCerts
    Dim objProduct
    Dim objSigner
    Dim iRet
    Dim arrGrants()
    Dim arrThumbprints
    Dim strThumbprint

    On Error Resume Next

    Set objSigner  = TkaGetSigner()
    Set objProduct = TkaGetProduct()

    iRet = objProduct.GetTokenActivationGrants(arrGrants)
    QuitIfError()
    
    arrThumbprints = objSigner.GetCertificateThumbprints(arrGrants)
    QuitIfError()
    
    For Each strThumbprint in arrThumbprints
        TkaPrintCertificate strThumbprint
    Next
End Sub


Private Sub TkaActivate(strThumbprint, strPin)
    Dim objService
    Dim objProduct
    Dim objSigner
    Dim iRet

    Dim strChallenge

    Dim strAuthInfo1
    Dim strAuthInfo2

    Set objSigner  = TkaGetSigner()
    Set objProduct = TkaGetProduct()
    Set objService = TkaGetService()

    DisplayActivatingSku objProduct

    On Error Resume Next

    iRet = objProduct.GenerateTokenActivationChallenge(strChallenge)
    QuitIfError()

    strAuthInfo1 = objSigner.Sign(strChallenge, strThumbprint, strPin, strAuthInfo2)
    QuitIfError()

    iRet = objProduct.DepositTokenActivationResponse(strChallenge, strAuthInfo1, strAuthInfo2)
    QuitIfError()

    objService.RefreshLicenseStatus()
    Err.Number = 0
    
    objProduct.refresh_
    DisplayActivatedStatus objProduct
    QuitIfError()

End Sub


Private Function TkaGetService()

    Set TkaGetService = GetServiceObject("Version")
    
End Function


Private Function TkaGetProduct()

    Dim objWinProductsWithPKeyInstalled
    Dim objProduct

    On Error Resume Next

    Set TkaGetProduct = Nothing

    Set TkaGetProduct = GetProductObject( _
                       "ID, Name, ApplicationId, PartialProductKey, Description, LicenseIsAddon ", _
                       "ApplicationId = '" & WindowsAppId & "' " &_
                       "AND PartialProductKey <> NULL " & _
                       "AND LicenseIsAddon = FALSE" _
                       )
    QuitIfError()

End Function

Private Function TkaGetSigner()

    On Error Resume Next
    Set TkaGetSigner = WScript.CreateObject("SPPWMI.SppWmiTokenActivationSigner")
    QuitIfError()

End Function

Private Sub TkaPrintCertificate(strThumbprint)
    Dim arrParams

    arrParams = Split(strThumbprint, "|")

    LineOut ""
    LineOut Replace(GetResource("L_MsgTkaCertThumbprint"), "%THUMBPRINT%", arrParams(0))
    LineOut Replace(GetResource("L_MsgTkaCertSubject"   ), "%SUBJECT%"   , arrParams(1))
    LineOut Replace(GetResource("L_MsgTkaCertIssuer"    ), "%ISSUER%"    , arrParams(2))
    LineOut Replace(GetResource("L_MsgTkaCertValidFrom" ), "%FROMDATE%"  , FormatDateTime(CDate(arrParams(3)), vbShortDate))
    LineOut Replace(GetResource("L_MsgTkaCertValidTo"   ), "%TODATE%"    , FormatDateTime(CDate(arrParams(4)), vbShortDate))
End Sub

' other generic options/helpers

Private Sub LineOut(str)
    g_EchoString = g_EchoString & str & vbNewLine
End Sub

Private Sub LineFlush(str)
    WScript.Echo g_EchoString & str
    g_EchoString = ""
End Sub

Private Sub ExitScript(retval)
    if (g_EchoString <> "") Then
        WScript.Echo g_EchoString
    End If
    WScript.Quit retval
End Sub

Private Sub InstallLicense(licFile)
    Dim objService
    Dim LicenseData
    Dim strOutput

    On Error Resume Next
    LicenseData = ReadAllTextFile(licFile)
    QuitIfError()
    set objService = GetServiceObject("Version")
    QuitIfError()

    objService.InstallLicense(LicenseData)
    QuitIfError()

    strOutput = Replace(GetResource("L_MsgLicenseFile"), "%LICENSEFILE%", licFile)
    LineOut strOutput
    LineOut ""
End Sub


' Returns the encoding for a givven file.
' Possible return values: ascii, unicode, unicodeFFFE (big-endian), utf-8
Function GetFileEncoding(strFileName)
    Dim strData
    Dim strEncoding
    Dim oStream

    Set oStream = CreateObject("ADODB.Stream")

    oStream.Type = 1 'adTypeBinary
    oStream.Open
    oStream.LoadFromFile(strFileName)

    ' Default encoding is ascii
    strEncoding =  "ascii"

    strData = BinaryToString(oStream.Read(2))

    ' Check for little endian (x86) unicode preamble
    If (Len(strData) = 2) and strData = (Chr(255) + Chr(254)) Then
        strEncoding = "unicode"
    Else
        oStream.Position = 0
        strData = BinaryToString(oStream.Read(3))

        ' Check for utf-8 preamble
        If (Len(strData) >= 3) and strData = (Chr(239) + Chr(187) + Chr(191)) Then
            strEncoding = "utf-8"
        End If
    End If

    oStream.Close

    GetFileEncoding = strEncoding
End Function

' Converts binary data (VT_UI1 | VT_ARRAY) to a string (BSTR)
Function BinaryToString(dataBinary)
  Dim i
  Dim str

  For i = 1 To LenB(dataBinary)
    str = str & Chr(AscB(MidB(dataBinary, i, 1)))
  Next

  BinaryToString = str
End Function

' Returns string containing the whole text file data.
' Supports ascii, unicode (little-endian) and utf-8 encoding.
Function ReadAllTextFile(strFileName)
    Dim strData
    Dim oStream
    
    Set oStream = CreateObject("ADODB.Stream")

    oStream.Type = 2 'adTypeText
    oStream.Open
    oStream.Charset = GetFileEncoding(strFileName)
    oStream.LoadFromFile(strFileName)

    strData = oStream.ReadText(-1) 'adReadAll

    oStream.Close

    ReadAllTextFile = strData
End Function

Private Function HandleOptionParam(cParam, mustProvide, opt, param)
    Dim strOutput

    HandleOptionParam = True
    If WScript.Arguments.Count <= cParam Then
        HandleOptionParam = False
        If mustProvide Then
            LineOut ""
            strOutput = Replace(GetResource("L_MsgErrorText_9"), "%OPTION%", opt)
            strOutput = Replace(strOutput, "%PARAM%", param)
            LineOut strOutput
            Call DisplayUsage()
        End If
    End If
End Function

Private Sub ShowErrorNum(strMessage, strErrNum)
    Dim strOutput
    Dim strErrText

    strErrText = GetResource("L_MsgError_" & strErrNum)
    if (strErrText = "") Then
        strOutput = Replace(GetResource("L_MsgErrorText_6"), "%ERRCODE%", strErrNum)
        LineOut strOutput
    End If

    LineOut strMessage & "0x" & strErrNum & " " & strErrText
End Sub

Private Sub QuitIfError()
    If Err.Number <> 0 Then
        ShowErrorNum GetResource("L_MsgErrorText_8"), Hex(Err.Number)
        ExitScript Err.Number
    End If
End Sub

Private Sub QuitWithError(errNum)
    ShowErrorNum GetResource("L_MsgErrorText_8"), CStr(errNum)
    ExitScript errNum
End Sub

Private Function booleanConnect()

    On Error Resume Next

    Dim objLocator, strOutput
    Dim objServer, objService
    Dim strErr, strVersion

    booleanConnect = True     'There is no error.

    'If this is the local computer, set everything and return immediately
    If g_strComputer = "." Then
        Set g_objWMIService = GetObject("winmgmts:\\" & g_strComputer & "\root\cimv2")
        Set g_objRegistry = GetObject("winmgmts:\\" & g_strComputer & "\root\default:StdRegProv")

        If Not g_serviceConnected Then
            g_serviceConnected = True
        End If
        Exit Function
    End If

    'Otherwise, establish the remote object connections

    ' Create Locator object to connect to remote CIM object manager
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    strErr = CStr(Hex(Err.Number))
    If Err.Number <> 0 Then
        strOutput = Replace(GetResource("L_MsgErrorWMI"), "%ERRCODE%", strErr)
        LineOut strOutput
        If Err.Description <> "" Then
            LineOut GetResource("L_MsgErrorDescription") & Err.Description & "."
        End If
        Err.Clear
        booleanConnect = False     'An error occurred
        Exit Function
    End If

    ' Connect to the namespace which is either local or remote
    Set g_objWMIService = objLocator.ConnectServer (g_strComputer, "\root\cimv2", g_strUserName, g_strPassword)
    strErr = CStr(Hex(Err.Number))
    If Err.Number <> 0 Then
        strOutput = Replace(GetResource("L_MsgErrorConnection"), "%ERRCODE%", strErr)
        strOutput = Replace(strOutput, "%COMPUTERNAME%", g_strComputer)
        LineOut strOutput
        If Err.Description <> "" Then
            LineOut GetResource("L_MsgErrorDescription") & Err.Description & "."
        End If
        Err.Clear
        booleanConnect = False     'An error occurred
        Exit Function
    End If

    g_objWMIService.Security_.impersonationlevel = wbemImpersonationLevelImpersonate
    strErr = CStr(Hex(Err.Number))
    If Err.Number <> 0 Then
        strOutput = Replace(GetResource("L_MsgErrorImpersonation"), "%ERRCODE%", strErr)
        LineOut strOutput
        If Err.Description <> "" Then
            LineOut GetResource("L_MsgErrorDescription") & Err.Description & "."
        End If
        Err.Clear
        booleanConnect = False     'An error occurred
        Exit Function
    End If

    g_objWMIService.Security_.AuthenticationLevel = wbemAuthenticationLevelPktPrivacy
    strErr = CStr(Hex(Err.Number))
    If Err.Number <> 0 Then
        strOutput = Replace(GetResource("L_MsgErrorAuthenticationLevel"), "%ERRCODE%", strErr)
        LineOut strOutput
        If Err.Description <> "" Then
            LineOut GetResource("L_MsgErrorDescription") & Err.Description & "."
        End If
        Err.Clear
        booleanConnect = False     'An error occurred
        Exit Function
    End If

    ' Get the SPP service version on the remote machine
    set objService = GetServiceObject("Version")
    strVersion = objService.Version
    
    ' The Windows 7 version of SLMgr.vbs does not support remote connections to Vista/WS08 machines
    if (Not IsNull(strVersion)) Then
        strVersion = Left(strVersion, 3)
        If (strVersion = "6.0") Then
            LineOut GetResource("L_MsgRemoteWmiVersionMismatch")
            booleanConnect = False
            Exit Function
        End If
    End If

    Set objServer = objLocator.ConnectServer(g_strComputer, "\root\default:StdRegProv", g_strUserName, g_strPassword)
    strErr = CStr(Hex(Err.Number))
    If Err.Number <> 0 Then
        strOutput = Replace(GetResource("L_MsgErrorConnectionRegistry"), "%ERRCODE%", strErr)
        strOutput = Replace(strOutput, "%COMPUTERNAME%", g_strComputer)
        LineOut strOutput
        If Err.Description <> "" Then
            LineOut GetResource("L_MsgErrorDescription") & Err.Description & "."
        End If
        Err.Clear
        booleanConnect = False     'An error occurred
        Exit Function
    End If

    objServer.Security_.ImpersonationLevel = 3
    Set g_objRegistry = objServer.Get("StdRegProv")
    strErr = CStr(Hex(Err.Number))
    If Err.Number <> 0 Then
        strOutput = Replace(GetResource("L_MsgErrorConnectionRegistry"), "%ERRCODE%", strErr)
        strOutput = Replace(strOutput, "%COMPUTERNAME%", g_strComputer)
        LineOut strOutput
        If Err.Description <> "" Then
            LineOut GetResource("L_MsgErrorDescription") & Err.Description & "."
        End If
        Err.Clear
        booleanConnect = False     'An error occurred
        Exit Function
    End If
    
    If Not g_serviceConnected Then
        g_serviceConnected = True
    End If
End Function

Function GetServiceObject(strQuery)
    Dim objService
    Dim colServices

    On Error Resume Next
       
    Set colServices = g_objWMIService.ExecQuery("SELECT " & strQuery & " FROM " & ServiceClass)
    QuitIfError()

    For each objService in colServices
        QuitIfError()
        Exit For
    Next

    set GetServiceObject = objService    
End Function

Function GetProductCollection(strSelect, strWhere)
    Dim colProducts

    On Error Resume Next

    If strWhere = EmptyWhereClause Then
        Set colProducts = g_objWMIService.ExecQuery("SELECT " & strSelect & " FROM " & ProductClass)
        QuitIfError()
    Else
        Set colProducts = g_objWMIService.ExecQuery("SELECT " & strSelect & " FROM " & ProductClass & " WHERE " & strWhere)
        QuitIfError()
        End If

    set GetProductCollection = colProducts    
End Function

Function GetProductObject(strSelect, strWhere)
    Dim objProduct
    Dim colProducts

    On Error Resume Next
    
    Set colProducts = GetProductCollection(strSelect, strWhere)
    QuitIfError()

    For each objProduct in colProducts
        QuitIfError()
        Exit For
    Next    

    set GetProductObject = objProduct    
End Function

Private Function IsKmsClient(strDescription)
    If InStr(strDescription, "VOLUME_KMSCLIENT") > 0 Then
        IsKmsClient = True
    Else
        IsKmsClient = False
    End If
End Function

Private Function  IsTkaClient(strDescription)
    IsTkaClient = IsKmsClient(strDescription)
End Function

Private Function IsKmsServer(strDescription)
    If IsKmsClient(strDescription) Then
        IsKmsServer = False
    Else
        If InStr(strDescription, "VOLUME_KMS") > 0 Then
            IsKmsServer = True
        Else
            IsKmsServer = False
        End If
    End If
End Function

Private Function IsTBL(strDescription)
    If InStr(strDescription, "TIMEBASED_") > 0 Then
        IsTBL = True
    Else
        IsTBL = False
    End If    
End Function

'Returns 0 if this is not the primary SKU, 1 if it is, and 2 if we aren't certain (older clients)
Function GetIsPrimaryWindowsSKU(objProduct)
    Dim iPrimarySku
    Dim bIsAddOn

    'Assume this is not the primary SKU
    iPrimarySku = 0
    'Verify the license is for Windows, that it has a partial key, and that
    If (LCase(objProduct.ApplicationId) = WindowsAppId And objProduct.PartialProductKey <> "") Then
        'If we can get verify the AddOn property then we can be certain
        On Error Resume Next
        bIsAddOn = objProduct.LicenseIsAddon
        If Err.Number = 0 Then
            If bIsAddOn = true Then
                iPrimarySku = 0
            Else
                iPrimarySku = 1
            End If
        Else
            'If we can not get the AddOn property then we assume this is a previous version
            'and we return a value of Uncertain, unless we can prove otherwise
            If (IsKmsClient(objProduct.Description) Or IsKmsServer(objProduct.Description)) Then
                'If the description is KMS related, we can be certain that this is a primary SKU
                iPrimarySku = 1
            Else
                'Indeterminate since the property was missing and we can't verify KMS
                iPrimarySku = 2
            End If
        End If
    End If
    GetIsPrimaryWindowsSKU = iPrimarySku
End Function

Private Function WasPrimaryKeyFound(strPrimarySkuType)
    If (IsKmsServer(strPrimarySkuType) Or IsKmsClient(strPrimarySkuType) Or (InStr(strPrimarySkuType, NotSpecialCasePrimaryKey) > 0) Or (InStr(strPrimarySkuType, TblPrimaryKey) > 0) Or (InStr(strPrimarySkuType, IndeterminatePrimaryKeyFound) > 0)) Then
        WasPrimaryKeyFound = True
    Else
        WasPrimaryKeyFound = False
    End If
End Function


Private Function CanPrimaryKeyTypeBeDetermined(strPrimarySkuType)
    If ((InStr(strPrimarySkuType, IndeterminatePrimaryKeyFound) > 0) Or (InStr(strPrimarySkuType, NoPrimaryKeyFound) > 0)) Then
        CanPrimaryKeyTypeBeDetermined = False
    Else
        CanPrimaryKeyTypeBeDetermined = True
    End If
End Function


Private Function GetPrimarySKUType()
    Dim objProduct
    Dim strPrimarySKUType, strDescription
    Dim iIsPrimaryWindowsSku

    For Each objProduct in GetProductCollection(ProductIsPrimarySkuSelectClause, PartialProductKeyNonNullWhereClause)
        strDescription = objProduct.Description
        If (LCase(objProduct.ApplicationId) = WindowsAppId) Then
            iIsPrimaryWindowsSku = GetIsPrimaryWindowsSKU(objProduct)
            If (iIsPrimaryWindowsSku = 1) Then
                If (IsKmsServer(strDescription) Or IsKmsClient(strDescription)) Then
                    strPrimarySKUType = strDescription
                    Exit For    'no need to continue
                Else
                    If IsTBL(strDescription) Then
                        strPrimarySKUType = TblPrimaryKey
                        Exit For
                    Else
                        strPrimarySKUType = NotSpecialCasePrimaryKey
                    End If
                End If
            ElseIf ((iIsPrimaryWindowsSku = 2) And strPrimarySKUType = "") Then
                strPrimarySKUType = IndeterminatePrimaryKeyFound
            End If
        Else
            strPrimarySKUType = strDescription
            Exit For    'no need to continue
        End If
    Next
    
    If strPrimarySKUType = "" Then
        strPrimarySKUType = NoPrimaryKeyFound
    End If

    GetPrimarySKUType = strPrimarySKUType
End Function

Private Function SetRegistryStr(hKey, strKeyPath, strValueName, strValue)
    SetRegistryStr = g_objRegistry.SetStringValue(hKey, strKeyPath, strValueName, strValue)
End Function

Private Function DeleteRegistryValue(hKey, strKeyPath, strValueName)
    DeleteRegistryValue = g_objRegistry.DeleteValue(hKey, strKeyPath, strValueName)
End Function

Private Function ExistsRegistryKey(hKey, strKeyPath)
    Dim bGranted
    Dim lRet

    ' Check for KEY_QUERY_VALUE for this key
    lRet = g_objRegistry.CheckAccess(hKey, strKeyPath, 1, bGranted)

    ' Ignore real access rights, just look for existence of the key
    If lRet<>2 Then
        ExistsRegistryKey = True
    Else
        ExistsRegistryKey = False
    End If
End Function

' Resource manipulation

' Get the resource string with the given name from the locale specific
' dictionary. If not found, use the built-in default.
Private Function GetResource(name)
    LoadResourceData
    If g_resourceDictionary.Exists(LCase(name)) Then
        GetResource = g_resourceDictionary.Item(LCase(name))
    Else
        GetResource = Eval(name)
    End If
End Function

' Loads resource strings from an ini file of the appropriate locale
Private Function LoadResourceData
    If g_resourcesLoaded Then
        Exit Function
    End If

    Dim ini, lang
    Dim fso
    
    Set fso = WScript.CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    lang = GetUILanguage()
    If Err.Number <> 0 Then
        'API does not exist prior to Vista so no resources to load
        g_resourcesLoaded = True
        Exit Function
    End If

    ini = fso.GetParentFolderName(WScript.ScriptFullName) & "\slmgr\" _
        & ToHex(lang) & "\" & fso.GetBaseName(WScript.ScriptName) &  ".ini"

    If fso.FileExists(ini) Then
        Dim stream
        Const ForReading = 1, TristateTrue = -1 'Read file in unicode format

        Set stream = fso.OpenTextFile(ini, ForReading, False, TristateTrue)
        ReadResources(stream)
        stream.Close
    End If

    g_resourcesLoaded = True
End Function

' Reads resource strings from an ini file
Private Function ReadResources(stream)
    const ERROR_FILE_NOT_FOUND = 2
    Dim ln, arr, key, value

    If Not IsObject(stream) Then Err.Raise ERROR_FILE_NOT_FOUND

    Do Until stream.AtEndOfStream
        ln = stream.ReadLine

        arr = Split(ln, "=", 2, 1)
        If UBound(arr, 1) = 1 Then
            ' Trim the key and the value first before trimming quotes
            key = LCase(Trim(arr(0)))
            value = TrimChar(Trim(arr(1)), """")

            If key <> "" Then
                g_resourceDictionary.Add key, value
            End If
        End If
    Loop
End Function

' Trim a character from the text string
Private Function TrimChar(s, c)
    Const vbTextCompare = 1

    ' Trim character from the start
    If InStr(1, s, c, vbTextCompare) = 1 Then
        s = Mid(s, 2)
    End If

    ' Trim character from the end
    If InStr(Len(s), s, c, vbTextCompare) = Len(s) Then
        s = Mid(s, 1, Len(s) - 1)
    End If

    TrimChar = s
End Function

' Get a 4-digit hexadecimal number
Private Function ToHex(n)
    Dim s : s = Hex(n)
    ToHex = String(4 - Len(s), "0") & s
End Function

