// This PAC file will provide proxy config to Microsoft 365 services
//  using data from the public web service for all endpoints
function FindProxyForURL(url, host)
{
    var direct = "DIRECT";
    var proxyServer = "PROXY 10.10.10.10:8080";



    if(shExpMatch(host, "*-files.sharepoint.com")
        || shExpMatch(host, "*-myfiles.sharepoint.com")
        || shExpMatch(host, "excelcs.officeapps.live.com")
        || shExpMatch(host, "ocws.officeapps.live.com")
        || shExpMatch(host, "odc.officeapps.live.com")
        || shExpMatch(host, "pptcs.officeapps.live.com")
        || shExpMatch(host, "roaming.officeapps.live.com")
        || shExpMatch(host, "uci.officeapps.live.com")
        || shExpMatch(host, "wordcs.officeapps.live.com"))
    {
        return proxyServer;
    }

    if(shExpMatch(host, "*.broadcast.skype.com")
        || shExpMatch(host, "*.compliance.microsoft.com")
        || shExpMatch(host, "*.lync.com")
        || shExpMatch(host, "*.mail.protection.outlook.com")
        || shExpMatch(host, "*.manage.office.com")
        || shExpMatch(host, "*.msftidentity.com")
        || shExpMatch(host, "*.msidentity.com")
        || shExpMatch(host, "*.officeapps.live.com")
        || shExpMatch(host, "*.online.office.com")
        || shExpMatch(host, "*.outlook.office.com")
        || shExpMatch(host, "*.portal.cloudappsecurity.com")
        || shExpMatch(host, "*.protection.office.com")
        || shExpMatch(host, "*.protection.outlook.com")
        || shExpMatch(host, "*.security.microsoft.com")
        || shExpMatch(host, "*.sharepoint.com")
        || shExpMatch(host, "*.skypeforbusiness.com")
        || shExpMatch(host, "*.teams.microsoft.com")
        || shExpMatch(host, "account.activedirectory.windowsazure.com")
        || shExpMatch(host, "account.office.net")
        || shExpMatch(host, "accounts.accesscontrol.windows.net")
        || shExpMatch(host, "admin.microsoft.com")
        || shExpMatch(host, "adminwebservice.microsoftonline.com")
        || shExpMatch(host, "api.passwordreset.microsoftonline.com")
        || shExpMatch(host, "autologon.microsoftazuread-sso.com")
        || shExpMatch(host, "becws.microsoftonline.com")
        || shExpMatch(host, "broadcast.skype.com")
        || shExpMatch(host, "clientconfig.microsoftonline-p.net")
        || shExpMatch(host, "companymanager.microsoftonline.com")
        || shExpMatch(host, "compliance.microsoft.com")
        || shExpMatch(host, "device.login.microsoftonline.com")
        || shExpMatch(host, "graph.microsoft.com")
        || shExpMatch(host, "graph.windows.net")
        || shExpMatch(host, "home.office.com")
        || shExpMatch(host, "login-us.microsoftonline.com")
        || shExpMatch(host, "login.microsoft.com")
        || shExpMatch(host, "login.microsoftonline-p.com")
        || shExpMatch(host, "login.microsoftonline.com")
        || shExpMatch(host, "login.windows.net")
        || shExpMatch(host, "logincert.microsoftonline.com")
        || shExpMatch(host, "loginex.microsoftonline.com")
        || shExpMatch(host, "manage.office.com")
        || shExpMatch(host, "nexus.microsoftonline-p.com")
        || shExpMatch(host, "office.live.com")
        || shExpMatch(host, "outlook.office.com")
        || shExpMatch(host, "outlook.office365.com")
        || shExpMatch(host, "passwordreset.microsoftonline.com")
        || shExpMatch(host, "portal.microsoftonline.com")
        || shExpMatch(host, "portal.office.com")
        || shExpMatch(host, "protection.office.com")
        || shExpMatch(host, "provisioningapi.microsoftonline.com")
        || shExpMatch(host, "security.microsoft.com")
        || shExpMatch(host, "smtp.office365.com")
        || shExpMatch(host, "teams.microsoft.com")
        || shExpMatch(host, "www.office.com"))
    {
        return direct;
    }

    return proxyServer;
}
