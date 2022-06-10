function FindProxyForURL(url,host)
{
// proxy-header v2.0

    url = url.toLowerCase();
    host = host.toLowerCase();

    // ============================================
    // SECTION: China Country Domains-->CN3B ZSCALER
    // ============================================

    if (
        shExpMatch(host, "*.cn")
        || shExpMatch(host, "*.com.cn")
        )
    return "PROXY 10.102.102.2:80";

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
        
https://www.google.com/maps/@31.2249757,121.4436474,18.5z
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

    // proxy-middle v2.0

    // ============================================
    // SECTION: IP Subnets that don't go to Zscaler
    // ============================================

    // Host is a L'Oreal IP address --> DIRECT
    // ---------------------------------------

    reip = /^\d+\.\d+\.\d+\.\d+$/;
    if (reip.test(host))

    return "PROXY 10.111.252.33:80"; 
} 
