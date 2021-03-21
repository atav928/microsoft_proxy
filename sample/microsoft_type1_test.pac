// This PAC file will provide proxy config to Microsoft 365 services
//  using data from the public web service for all endpoints
function FindProxyForURL(url, host)
{
    var direct = "DIRECT";
    var proxyServer = "PROXY 10.10.10.10:8080";



    if(shExpMatch(host, "*-files.sharepoint.com")
        || shExpMatch(host, "*-myfiles.sharepoint.com"))
    {
        return proxyServer;
    }

    if(shExpMatch(host, "*.sharepoint.com")
        || shExpMatch(host, "outlook.office.com")
        || shExpMatch(host, "outlook.office365.com"))
    {
        return direct;
    }

    return proxyServer;
}
