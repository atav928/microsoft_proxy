// This PAC file will provide proxy config to Microsoft 365 services
//  using data from the public web service for all endpoints
function FindProxyForURL(url, host)
{
    var direct = "DIRECT";
    var proxyServer = "PROXY http://proxy-sav.gac.gulfaero.com:8080; http://proxy-atl.gac.gulfaero.com:8080; direct";





    if(shExpMatch(host, "gulfaero-my.sharepoint.com")
        || shExpMatch(host, "gulfaero.sharepoint.com")
        || shExpMatch(host, "outlook.office.com")
        || shExpMatch(host, "outlook.office365.com"))
    {
        return direct;
    }

    return proxyServer;
}
