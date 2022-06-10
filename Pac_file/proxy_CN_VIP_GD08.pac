function FindProxyForURL(url,host)
{
// proxy-header v2.0

    url = url.toLowerCase();
    host = host.toLowerCase();

    // =======================================
    // SECTION: Hosts that don't go to Zscaler
    // =======================================

    // Hosts handled by --> EMEA ITRA proxy
    // ------------------------------------

    if (
        shExpMatch(host,"mdmadmin.loreal.net")
        || shExpMatch(host,"www.adpworld.de")
        || shExpMatch(host,"*.patentorder.com")
        || shExpMatch(host,"sts-ri.loreal.net")
        || shExpMatch(host,"*.sendgrid.com")
        || shExpMatch(host,"*.gcsip.com")
        || shExpMatch(host,"*.gcsip.nl")
        )
    return "PROXY 10.3.0.1:8080";

    // Hosts handled by --> SGD ITRA proxy
    // -----------------------------------

    if (
        dnsDomainIs(host,"www.ctrlpay.com")
        )
    return "PROXY 10.119.252.33:80";

    // Hosts handled by --> US ITRA Proxy
    // ----------------------------------

    if (
        shExpMatch(host,"plcaching.na.loreal.intra")
        || shExpMatch(host,"*.kroger.com")
        || shExpMatch(host,"sendfile.lorealamericas.com")
        || shExpMatch(host,"icommerceagent.pfsweb.com")
        || shExpMatch(host,"*loginbox.loreal.net")
        || shExpMatch(host,"services.loreal.ca")
        || shExpMatch(host,"*.lorealusa.com")
        || shExpMatch(host,"analyst.at-hand.net")
        || shExpMatch(host,"analyst2.at-hand.net")
        || shExpMatch(host,"*.compusensecloud.com")
        || shExpMatch(host,"walgreens-fm.predictix.com")
        )
    return "PROXY 10.24.148.12:80";

    // Internal domains --> DIRECT
    // ---------------------------

    if (
        isPlainHostName(host)
        || shExpMatch(host,"*.local")
        || shExpMatch(host,"*.loreal.intra")
        || shExpMatch(host,"*.loreal.wans")
        || shExpMatch(host,"*.lorealitalia.intra")
        || shExpMatch(host,"*.lorealamericas.com")
        || shExpMatch(host,"*.rd.loreal")
        )
    return "DIRECT";

    // Internal view loreal.com --> DIRECT
    // -----------------------------------

    if (
        shExpMatch(host,"*.us.loreal.com")
        || shExpMatch(host,"*webconf.loreal.com")
        || shExpMatch(host,"*autodiscover.loreal.com")
        || shExpMatch(host,"autodiscover.rd.loreal.com")
        || shExpMatch(host,"*mail.loreal.com")
        || shExpMatch(host,"*mail-ap.loreal.com")
        || shExpMatch(host,"*mail-na.loreal.com")
        || shExpMatch(host,"webconf-na.loreal.com")
        || shExpMatch(host,"*discoverinternal.loreal.com")
        )
    return "DIRECT";

    // loreal.net with internal IP address --> DIRECT
    // ---------------------------------------------

    if (shExpMatch(host,"*.loreal.net")
        // Excluding dual-view hosts that do not work on their internal IP
        && !shExpMatch(host,"dev-marsapins-apac.loreal.net")
        && !shExpMatch(host,"dev-marsmobilityns-apac.loreal.net")
        && !shExpMatch(host,"dev-marsns-apac.loreal.net")
        && !shExpMatch(host,"hrkpi-cn.loreal.net")
        && !shExpMatch(host,"marsapins-apac.loreal.net")
        && !shExpMatch(host,"marsas-apac.loreal.net")
        && !shExpMatch(host,"marsasdev-apac.loreal.net")
        && !shExpMatch(host,"marsasqa-apac.loreal.net")
        && !shExpMatch(host,"marsmobility-cn.loreal.net")
        && !shExpMatch(host,"marsmobilitydev-cn.loreal.net")
        && !shExpMatch(host,"marsmobilityns-apac.loreal.net")
        && !shExpMatch(host,"marsns-apac.loreal.net")
        && !shExpMatch(host,"preprod-marsapins-apac.loreal.net")
        && !shExpMatch(host,"preprod-mars-hk.loreal.net")
        && !shExpMatch(host,"preprod-marsmobilityns-apac.loreal.net")
        && !shExpMatch(host,"preprod-marsns-apac.loreal.net")
        && !shExpMatch(host,"preprod-primenews.loreal.net")
        && !shExpMatch(host,"qua-onemarket.loreal.net")
        && !shExpMatch(host,"starnet-q.loreal.net")
        && !shExpMatch(host,"click-tw.loreal.net")
        )
        {
        var resolved_ip = dnsResolve(host);
        if (
            isInNet(resolved_ip, "10.0.0.0", "255.0.0.0")
            || isInNet(resolved_ip, "172.16.0.0", "255.240.0.0")
            || isInNet(resolved_ip, "192.168.0.0", "255.255.0.0")
            || isInNet(resolved_ip, "127.0.0.0", "255.0.0.0")
            || isInNet(resolved_ip, "169.254.0.0", "255.255.0.0")
            || isInNet(resolved_ip, "22.0.0.0", "255.0.0.0")
            || isInNet(resolved_ip, "29.0.0.0", "255.0.0.0")
            || isInNet(resolved_ip, "30.0.0.0", "255.0.0.0")
            || isInNet(resolved_ip, "192.6.0.0", "255.254.0.0")
            || isInNet(resolved_ip, "192.8.0.0", "255.255.0.0")
            || isInNet(resolved_ip, "128.0.0.0", "255.0.0.0")
            || isInNet(resolved_ip, "12.1.1.10", "255.255.255.255")
            || isInNet(resolved_ip, "161.241.0.0", "255.255.0.0")
            || isInNet(resolved_ip, "200.200.200.0", "255.255.254.0")
            )
        return "DIRECT";
        }

    // Special hosts --> DIRECT
    // ------------------------

    if (
        shExpMatch(host,"*infra.oa.acn")
        || shExpMatch(host,"*lorealts.nordman.se")
        || shExpMatch(host,"*lorealdamcs.cloudapp.net")
        || shExpMatch(host,"*oadamcs.cloudapp.net")
        || shExpMatch(host,"*edit.boxlocalhost.com")
        || shExpMatch(host,"*.regefi.com")
        || shExpMatch(host,"*.pacbiolabs.com")
        || shExpMatch(host,"*.internalclarisonic.com")
        || shExpMatch(host,"*.clarisonic.biz")
        || shExpMatch(host,"*.kapayroll.com")
        || shExpMatch(host,"*.ccc.cea.fr")
        || shExpMatch(host,"vpn.sephora.fr")
        )
    return "DIRECT";

// CN VIP specific

    if (
        shExpMatch(host, "*.cn")
        || shExpMatch(host, "*.com.cn")
        )
    return "DIRECT";

    // CHINA Site Exceptions for OFFICE365
    if (
        // Teams live events
        shExpMatch(host,"*.teams.microsoft.com")
        || shExpMatch(host,"*.media.azure.net")
        || shExpMatch(host,"*.streaming.mediaservices.windows.net")
        || shExpMatch(host,"*.hivestreaming.com")

        // Power Bi
        || shExpMatch(host,"*.powerbi.com")
        || shExpMatch(host,"*.powerapps.com")
        || shExpMatch(host,"wabi-north-europe-d-primary-redirect.analysis.windows.net")
        || shExpMatch(host,"loreal-myfiles.sharepoint.com")
        || shExpMatch(host,"*.msauth.net")
        )
    return "PROXY 185.46.212.89:80";

    if (
    // Sharepoint, OneDrive official URLs
        shExpMatch(host,"loreal.sharepoint.com")
        || shExpMatch(host,"loreal-my.sharepoint.com")
        || shExpMatch(host,"*.log.optimizely.com")
        || shExpMatch(host,"ssw.live.com")
        || shExpMatch(host,"storage.live.com")
        || shExpMatch(host,"*.wns.windows.com")
        || shExpMatch(host,"admin.onedrive.com")
        || shExpMatch(host,"officeclient.microsoft.com")
        || shExpMatch(host,"g.live.com")
        || shExpMatch(host,"oneclient.sfx.ms")
        || shExpMatch(host,"*.svc.ms")
        || shExpMatch(host,"loreal-files.sharepoint.com")
        )
    return "PROXY 185.46.212.89:80";

    // O365 URLs via HKT
    if (
        // Aggreg of official O365 URLs, to lower number of rules
        shExpMatch(host,"*.office365.com")
        || shExpMatch(host,"*.office.com")
        || shExpMatch(host,"*.office.net")
        || shExpMatch(host,"*.officeapps.live.com")
        || shExpMatch(host,"*.outlook.com")
        || shExpMatch(host,"*.microsoftonline.com")
        || shExpMatch(host,"*.microsoftonline-p.com")
        || shExpMatch(host,"*.microsoftonline-p.net")
        || shExpMatch(host,"*.msecnd.net")
        || shExpMatch(host,"*.msftncsi.com")
        || shExpMatch(host,"*.trafficmanager.net")
        || shExpMatch(host,"*.sharepointonline.com")
        || shExpMatch(host,"*.azureedge.net")
        || shExpMatch(host,"*.keydelivery.mediaservices.windows.net")
        // || shExpMatch(host,"*.streaming.mediaservices.windows.net")
        || shExpMatch(host,"*.blob.core.windows.net")
        || shExpMatch(host,"*.msocdn.com")
        || shExpMatch(host,"*.assets-yammer.com")
        || shExpMatch(host,"ajax.aspnetcdn.com")
        || shExpMatch(host,"watson.telemetry.microsoft.com")

        // Skype, Teams official URLs
        || shExpMatch(host,"*.lync.com")
        // || shExpMatch(host,"*.teams.microsoft.com")
        || shExpMatch(host,"teams.microsoft.com")
        || shExpMatch(host,"*.sfbassets.com")
        || shExpMatch(host,"*.urlp.sfbassets.com")
        || shExpMatch(host,"aka.ms")
        || shExpMatch(host,"amp.azure.net")
        || shExpMatch(host,"*.users.storage.live.com")
        || shExpMatch(host,"*.adl.windows.com")
        || shExpMatch(host,"*.msedge.net")
        || shExpMatch(host,"compass-ssl.microsoft.com")
        || shExpMatch(host,"*.mstea.ms")
        || shExpMatch(host,"*.secure.skypeassets.com")
        || shExpMatch(host,"*.skype.com")
        || shExpMatch(host,"*.skypeforbusiness.com")
        || shExpMatch(host,"*.tenor.com")
        || shExpMatch(host,"statics.teams.microsoft.com")

        // Extra URLs 
        || shExpMatch(host,"vortex.data.microsoft.com")
        || shExpMatch(host,"*.aria.microsoft.com")
        || shExpMatch(host,"*.events.data.microsoft.com")
        || shExpMatch(host,"ecs.office.com")
        || shExpMatch(host,"teams.live.com")
        || shExpMatch(host,"lyncdiscover.lntinfotech.com")
        || shExpMatch(host,"sip.lntinfotech.com")
        || shExpMatch(host,"v10.vortex-win.data.microsoft.com")
        || shExpMatch(host,"statics.teams.cdn.office.net")

        // Exchange official URLs
        || shExpMatch(host,"outlook.office.com")
        || shExpMatch(host,"outlook.office365.com")
        || shExpMatch(host,"r1.res.office365.com")
        || shExpMatch(host,"r3.res.office365.com")
        || shExpMatch(host,"r4.res.office365.com")
        || shExpMatch(host,"*.outlook.com")
        || shExpMatch(host,"*.outlook.office.com")
        || shExpMatch(host,"attachments.office.net")
        || shExpMatch(host,"*.protection.outlook.com")
        || shExpMatch(host,"autodiscover.loreal.onmicrosoft.com")
        || shExpMatch(host,"*.azuredege.net")
        || shExpMatch(host,"web.microsoftstream.com")

        // Sharepoint, OneDrive official URLs
       // || shExpMatch(host,"loreal.sharepoint.com")
       // || shExpMatch(host,"loreal-my.sharepoint.com")
       // || shExpMatch(host,"*.log.optimizely.com")
       // || shExpMatch(host,"ssw.live.com")
       // || shExpMatch(host,"storage.live.com")
       // || shExpMatch(host,"*.wns.windows.com")
       // || shExpMatch(host,"admin.onedrive.com")
       // || shExpMatch(host,"officeclient.microsoft.com")
       // || shExpMatch(host,"g.live.com")
       // || shExpMatch(host,"oneclient.sfx.ms")
       // || shExpMatch(host,"*.svc.ms")
       // || shExpMatch(host,"loreal-files.sharepoint.com")

        // O365 Common and Office official URLs
        || shExpMatch(host,"*.api.microsoftstream.com")
        || shExpMatch(host,"*.notification.api.microsoftstream.com")
        || shExpMatch(host,"api.microsoftstream.com")
        || shExpMatch(host,"web.microsoftstream.com")
        || shExpMatch(host,"nps.onyx.azure.net")
        // || shExpMatch(host,"*.media.azure.net")
        || shExpMatch(host,"office.live.com")
        || shExpMatch(host,"*.onenote.com")
        || shExpMatch(host,"*cdn.onenote.net")
        || shExpMatch(host,"apis.live.net")
        || shExpMatch(host,"cdn.optimizely.com")
        || shExpMatch(host,"officeapps.live.com")
        || shExpMatch(host,"www.onedrive.com")
        || shExpMatch(host,"*.msappproxy.net")
        || shExpMatch(host,"*.msftidentity.com")
        || shExpMatch(host,"*.msidentity.com")
        || shExpMatch(host,"account.activedirectory.windowsazure.com")
        || shExpMatch(host,"accounts.accesscontrol.windows.net")
        || shExpMatch(host,"autologon.microsoftazuread-sso.com")
        || shExpMatch(host,"graph.microsoft.com")
        || shExpMatch(host,"graph.windows.net")
        || shExpMatch(host,"login.microsoft.com")
        || shExpMatch(host,"login.windows.net")
        || shExpMatch(host,"*.msauth.net")
        || shExpMatch(host,"*.msauthimages.net")
        || shExpMatch(host,"*.msftauth.net")
        || shExpMatch(host,"*.msftauthimages.net")
        || shExpMatch(host,"*.phonefactor.net")
        || shExpMatch(host,"enterpriseregistration.windows.net")
        || shExpMatch(host,"management.azure.com")
        || shExpMatch(host,"policykeyservice.dc.ad.msft.net")
        || shExpMatch(host,"*.portal.cloudappsecurity.com")
        || shExpMatch(host,"admin.microsoft.com")
        || shExpMatch(host,"*.o365weve.com")
        || shExpMatch(host,"appsforoffice.microsoft.com")
        || shExpMatch(host,"assets.onestore.ms")
        || shExpMatch(host,"auth.gfx.ms")
        || shExpMatch(host,"c1.microsoft.com")
        || shExpMatch(host,"client.hip.live.com")
        || shExpMatch(host,"dgps.support.microsoft.com")
        || shExpMatch(host,"docs.microsoft.com")
        || shExpMatch(host,"msdn.microsoft.com")
        || shExpMatch(host,"platform.linkedin.com")
        || shExpMatch(host,"support.microsoft.com")
        || shExpMatch(host,"technet.microsoft.com")
        || shExpMatch(host,"*.cloudapp.net")
        || shExpMatch(host,"*.aadrm.com")
        || shExpMatch(host,"*.azurerms.com")
        || shExpMatch(host,"*.informationprotection.azure.com")
        || shExpMatch(host,"ecn.dev.virtualearth.net")
        || shExpMatch(host,"informationprotection.hosting.portal.azure.net")
        || shExpMatch(host,"testconnectivity.microsoft.com")
        || shExpMatch(host,"*.hockeyapp.net")
        || shExpMatch(host,"dc.applicationinsights.microsoft.com")
        || shExpMatch(host,"dc.services.visualstudio.com")
        || shExpMatch(host,"forms.microsoft.com")
        || shExpMatch(host,"mem.gfx.ms")
        || shExpMatch(host,"signup.microsoft.com")
        || shExpMatch(host,"staffhub.ms")
        || shExpMatch(host,"staffhub.uservoice.com")
        || shExpMatch(host,"*.onmicrosoft.com")
        || shExpMatch(host,"o15.officeredir.microsoft.com")
        || shExpMatch(host,"officepreviewredir.microsoft.com")
        || shExpMatch(host,"officeredir.microsoft.com")
        || shExpMatch(host,"r.office.microsoft.com")
        || shExpMatch(host,"activation.sls.microsoft.com")
        || shExpMatch(host,"crl.microsoft.com")
        || shExpMatch(host,"office15client.microsoft.com")
        || shExpMatch(host,"officeclient.microsoft.com")
        || shExpMatch(host,"go.microsoft.com")
        || shExpMatch(host,"officecdn.microsoft.com")
        || shExpMatch(host,"officecdn.microsoft.com.edgesuite.net")
        || shExpMatch(host,"*.yammer.com")
        || shExpMatch(host,"*.yammerusercontent.com")
        || shExpMatch(host,"eus-www.sway-cdn.com")
        || shExpMatch(host,"eus-www.sway-extensions.com")
        || shExpMatch(host,"wus-www.sway-cdn.com")
        || shExpMatch(host,"wus-www.sway-extensions.com")
        || shExpMatch(host,"sway.com")
        || shExpMatch(host,"www.sway.com")
        || shExpMatch(host,"*.manage.microsoft.com")
        || shExpMatch(host,"cdnprod.myanalytics.microsoft.com")
        || shExpMatch(host,"myanalytics.microsoft.com")
        || shExpMatch(host,"myanalytics-gcc.microsoft.com")
        )
    return "PROXY 185.46.212.88:80";

// proxy-middle v2.0

    // ============================================
    // SECTION: IP Subnets that don't go to Zscaler
    // ============================================

    // Host is a L'Oreal IP address --> DIRECT
    // ---------------------------------------

    reip = /^\d+\.\d+\.\d+\.\d+$/;
    if (reip.test(host))
    {
    if (
        isInNet(host, "10.0.0.0", "255.0.0.0")
        || isInNet(host, "22.0.0.0", "255.0.0.0")
        || isInNet(host, "29.0.0.0", "255.0.0.0")
        || isInNet(host, "30.0.0.0", "255.0.0.0")
        || isInNet(host, "192.168.0.0", "255.255.0.0")
        || isInNet(host, "192.6.0.0", "255.254.0.0")
        || isInNet(host, "192.8.0.0", "255.255.0.0")
        || isInNet(host, "127.0.0.0", "255.255.0.0")
        || isInNet(host, "128.0.0.0", "255.0.0.0")
        || isInNet(host, "172.16.0.0", "255.240.0.0")
        || isInNet(host, "12.1.1.10", "255.255.255.255")
        || isInNet(host, "161.241.0.0", "255.255.0.0")
        || isInNet(host, "200.200.200.0", "255.255.254.0")
        )
    return "DIRECT";
    }

    return "PROXY 10.243.160.73:80"; 
} 
