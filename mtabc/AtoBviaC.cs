﻿using System;
using System.Net;
using System.Text;
using System.Runtime.InteropServices;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;
using tankers.distances.models.datawindow;
using tankers.distances.models;
using System.Collections;
using System.Web;
using System.Collections.Specialized;
using System.Globalization;

/* 
 * ======
 * CR4343 
 * ======
 * created 20161205
 * last modified 20171101
 * 
 * leg index starts at 0
 * journey index starts at 1
 */
namespace tankers.distances
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("mt.distances.AtoBviaC")]
    
    public class AtoBviaC
    {
        private static XNamespace _xmlns;
        private static XNamespace _wsns;
        private String _apikey;
        private String _wsparmstring;
        private XDocument _routingpointxml_cached;
        public XDocument _voyagexml;
        private XDocument _portsxml_cached;
        private StringBuilder _rpShortCodesInLeg = new StringBuilder();
        private StringBuilder _rpNamesInLeg = new StringBuilder();
        private StringBuilder _rpOpenByDefaultInLeg = new StringBuilder();
        private StringBuilder _routingsForDW = new StringBuilder("");

        enum DataFormat
        {
            None,
            XML,
            PBDataWindowSyntax,
            wsParmString
        };

        /// <summary>
        /// the constructor of the main class AtoBviaC
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20161205
        /// last modified       20170119
        /// author              AGL027
        /// </remarks>
        public AtoBviaC()
        {
            _xmlns = "http://api.atobviaconline.com/v1";
            _wsns = "https://api.atobviaconline.com/v1";
            /* this method does not require api_key */
            _routingpointxml_cached = _getRoutingPoints();
        }

        /// <summary>
        /// Direct call to web service that gets the data for all Routing Points and allows us to load this into an
        /// XMLDocument data object.  This method is not dependent on a valid api_key.
        /// It is called from the constructor method.
        /// TODO - could be optimized
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20161206
        /// last modified       20170119
        /// author              AGL027
        /// </remarks>
        private XDocument _getRoutingPoints()
        {

            StringBuilder url = new StringBuilder();
            url.Append(_wsns).Append("/RoutingPoints?api_key=").Append(_apikey);

            WebClient syncClient = new WebClient();
            syncClient.Headers.Add("accept", "application/xml");
            String content = syncClient.DownloadString(url.ToString());
            XDocument rpXml = new XDocument();

            try
            {
                rpXml = XDocument.Parse(content);
            }
            catch (System.Xml.XmlException)
            {
                System.Windows.Forms.MessageBox.Show("error in obtaining routing point data");
            }
            return rpXml;
        }

        /// <summary>
        /// Direct call to web service that gets the data for all Ports and allows us to load this into an
        /// XMLDocument data object.  This method is dependent on a valid api_key.
        /// It is called from the constructor method.
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20170119
        /// last modified       20170119
        /// author              AGL027
        /// </remarks>
        private XDocument _getPorts()
        {
            StringBuilder url = new StringBuilder();
            url.Append(_wsns).Append("/Ports?api_key=").Append(_apikey);

            WebClient syncClient = new WebClient();
            syncClient.Headers.Add("accept", "application/xml");
            String content = syncClient.DownloadString(url.ToString());
            XDocument portXml = new XDocument();
            try
            {
                portXml = XDocument.Parse(content);
            }
            catch (System.Xml.XmlException)
            {
                System.Windows.Forms.MessageBox.Show("error in obtaining routing point data");
            }
            return portXml;
        }

        /// <summary>
        /// Simply get OpenByDefault value from the routing point XMLDocument object.
        /// </summary>
        /// 
        /// <remarks>
        /// date created        ??????
        /// last modified       20171101
        /// author              AGL027
        /// </remarks>
        private String _getOpenByDefault(String sRpShortCode)
        {
            var rps = (from rp in _routingpointxml_cached.Descendants(_xmlns + "RoutingPoint")
                       where rp.Element(_xmlns + "ShortCode").Value.Contains(sRpShortCode)
                       select rp.Element(_xmlns + "OpenByDefault").Value).FirstOrDefault();
            return rps;
        }


        /// <summary>
        /// Obtain delimited routing data from the voyage XML string
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20170116
        /// last modified       20170118
        /// author              AGL027
        /// </remarks>
        private StringBuilder _parseXMLContentToDWFormat()
        {
            ArrayList Routing = new ArrayList();

            StringBuilder wayPointData = new StringBuilder();
            StringBuilder fromPortData = new StringBuilder();
            StringBuilder toPortData = new StringBuilder();
            StringBuilder dwContent = new StringBuilder();

            int countOfLegs = (int)_voyagexml.Descendants(_xmlns + "Leg").Count();

            var legs = (from e in _voyagexml.Descendants(_xmlns + "Legs").Elements(_xmlns + "Leg")
                        select new Leg()
                        {

                            fromPort = e.Elements(_xmlns + "FromPort")
                            .Select(r => new Port()
                            {
                                name = (string)r.Element(_xmlns + "Name"),
                                code = (string)r.Element(_xmlns + "Code")

                            }).FirstOrDefault(),

                            toPort = e.Elements(_xmlns + "ToPort")
                            .Select(r => new Port()
                            {
                                name = r.Element(_xmlns + "Name") != null ? r.Element(_xmlns + "Name").Value : "",
                                code = r.Element(_xmlns + "Code") != null ? r.Element(_xmlns + "Code").Value : "",

                            }).First(),

                            WayPointList = e.Element(_xmlns + "Waypoints").Elements(_xmlns + "Waypoint")
                            .Select(r => new WayPoint()
                            {
                                name = r.Element(_xmlns + "Name") != null ? r.Element(_xmlns + "Name").Value : "",
                                DistanceFromStart = Convert.ToDecimal(r.Element(_xmlns + "DistanceFromStart").Value),
                                routingPointCode = r.Element(_xmlns + "RoutingPoint") != null ? r.Element(_xmlns + "RoutingPoint").Value : "",
                                EcaZoneToPrevious = r.Element(_xmlns + "EcaZoneToPrevious") != null ? r.Element(_xmlns + "EcaZoneToPrevious").Value : ""

                            }).ToList()

                        }).ToList();

            int legIndex = 0;
            foreach (Leg leg in legs)
            {
                legIndex++;
                decimal distancePorttoPort = 0;
                decimal sumOfDistancesSinceLastKnown = 0;
                decimal LastPortDistanceFromStart = 0;
                decimal sumOfEcaZoneDistancesSinceLastKnown = 0;

                wayPointData = new StringBuilder();
                fromPortData = new StringBuilder();
                toPortData = new StringBuilder();

                int wp_counter = 0;
                int is_eca = 0;

                foreach (WayPoint wp in leg.WayPointList)
                {
                    distancePorttoPort = wp.DistanceFromStart - LastPortDistanceFromStart;
                    sumOfDistancesSinceLastKnown += distancePorttoPort;

                    if (wp.EcaZoneToPrevious != "")
                    {
                        is_eca = 1;
                        sumOfEcaZoneDistancesSinceLastKnown += distancePorttoPort;
                    }
                    else
                    {
                        is_eca = 0;
                        sumOfEcaZoneDistancesSinceLastKnown = 0;
                    }

                    if (leg.fromPort.code == wp.name && legIndex == 1)
                    {
                        /* we are in the first port of the first leg - write out the initial port */
                        fromPortData.Append(leg.fromPort.code).Append("\t").Append(leg.fromPort.name).Append("\t").Append("0.000").Append("\t").Append("0.000").Append("\t").Append(is_eca).Append("\r\n");
                    }
                    else if (leg.toPort.code == wp.name)
                    {
                        /* we are in the last port of the leg.  We always write this data out. */
                        toPortData.Append(leg.toPort.code).Append("\t").Append(leg.toPort.name).Append("\t").Append(sumOfDistancesSinceLastKnown.ToString()).Append("\t").Append(sumOfEcaZoneDistancesSinceLastKnown.ToString()).Append("\t").Append(is_eca).Append("\r\n");
                    }
                    else
                    {
                        if ((wp.name.Length > 4 && wp.name.Substring(0, 5) == "Exit ") || wp.name.IndexOf("Enter ") >= 0)
                        {
                            wayPointData.Append("").Append("\tECA:").Append(wp.name).Append("\t").Append(sumOfDistancesSinceLastKnown.ToString()).Append("\t").Append(sumOfEcaZoneDistancesSinceLastKnown.ToString()).Append("\t").Append(is_eca).Append("\r\n");

                        }

                        else if (wp.routingPointCode != "")
                        {
                            wayPointData.Append(wp.routingPointCode).Append("\t").Append(GetRPDataByShortCode(wp.routingPointCode, 1)).Append("\t").Append(sumOfDistancesSinceLastKnown.ToString()).Append("\t").Append(sumOfEcaZoneDistancesSinceLastKnown.ToString()).Append("\t").Append(is_eca).Append("\r\n");
                            sumOfDistancesSinceLastKnown = 0;
                            sumOfEcaZoneDistancesSinceLastKnown = 0;
                        }

                        if (wp.name.Length > 4 && wp.name.Substring(0, 5) == "Exit " || wp.name.IndexOf("Enter ") >= 0)
                        {
                            sumOfEcaZoneDistancesSinceLastKnown = 0;
                            sumOfDistancesSinceLastKnown = 0;
                        }

                    }
                    LastPortDistanceFromStart = wp.DistanceFromStart;
                    wp_counter++;
                }
                dwContent.Append(fromPortData).Append(wayPointData).Append(toPortData);
            }
            return dwContent;
        }

        /// <summary>
        /// this must be called by the requestor to load the necessary api_key that is needed by AtoBviaC.
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20161215
        /// last modified       20171101
        /// author              AGL027
        /// </remarks>
        /// <param name="apikey">api_key must be stored in client application.</param> 
        public void SetApiKey(String sApiKey)
        {
            _apikey = sApiKey;
        }

        /// <summary>
        /// On the settings page, this gets the account details.
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20170530
        /// last modified       20170531
        /// author              AGL027
        /// </remarks>
        public String GetAccountDetails()
        {
            StringBuilder url = new StringBuilder();
            StringBuilder accountDetails = new StringBuilder("");

            url.Append(_wsns.ToString()).Append("/AccountDetails?").Append("api_key=" + _apikey);

            abcResponse abcResp = new abcResponse();
            abcResp = _downloadStringFromURL(url.ToString());

            if (!string.IsNullOrEmpty(abcResp.httpcode))
            {
                return abcResp.httpcode;
            }
            else
            {
                XElement accountElement = XElement.Parse(abcResp.content);

                string version = accountElement.Element(_xmlns + "Version").Value;
                string licenceexpiry = accountElement.Element(_xmlns + "LicenceExpiry").Value;
                string remainingdistances = accountElement.Element(_xmlns + "RemainingDistances").Value;

                accountDetails.Append(version).Append("|").Append(licenceexpiry).Append("|").Append(remainingdistances);
            }

            return accountDetails.ToString();
        }


        /// <summary>
        /// On the settings page, this gets all the routing points used.
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20170531
        /// last modified       20170531
        /// author              AGL027
        /// </remarks>
        public String GetAllRoutingsForDW()
        {
            // we already have routingpointxml data!
            StringBuilder routingList = new StringBuilder("");

            var routings = from rp in _routingpointxml_cached.Descendants(_xmlns + "RoutingPoint")
                           select new
                           {
                               shortcode = rp.Element(_xmlns + "ShortCode").Value.ToString(),
                               rpname = rp.Element(_xmlns + "Name").Value.ToString(),
                               openbydefault = rp.Element(_xmlns + "OpenByDefault").Value.ToString()
                           };

            foreach (var rp in routings)
            {
                routingList.Append(rp.shortcode).Append("\t").Append(rp.rpname).Append("\t");
                if (rp.openbydefault.ToUpper() == "TRUE")
                {
                    routingList.Append("1").Append("\r\n");
                }
                else
                {
                    routingList.Append("0").Append("\r\n");
                }
            }
            return routingList.ToString();
        }


        /// <summary>
        /// On the settings page, this gets all the ports used.
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20170531
        /// last modified       20170531
        /// author              AGL027
        /// </remarks>
        public String GetAllPortsForDW()
        {
            StringBuilder portList = new StringBuilder("");
            /* we might not have the port data cached yet*/
            if (_portsxml_cached == null)
            {
                _portsxml_cached = _getPorts();
            }

            var ports = from p in _portsxml_cached.Descendants(_xmlns + "Port")
                        select new
                        {
                            code = p.Element(_xmlns + "Code").Value.ToString(),
                            countrycode = p.Element(_xmlns + "CountryCode").Value.ToString(),
                            portname = p.Element(_xmlns + "Name").Value.ToString(),
                            lat = p.Element(_xmlns + "LatGeodetic").Value.ToString(),
                            lon = p.Element(_xmlns + "Lon").Value.ToString()
                        };

            foreach (var p in ports)
            {
                portList.Append(p.code).Append("\t").Append(p.countrycode).Append("\t").Append(p.portname).Append("\t").Append(p.lat).Append("\t").Append(p.lon).Append("\r\n");

            }
            return portList.ToString();
        }



        /// <summary>
        /// low level communication to web service; this is to handle strings
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20161205
        /// last modified       20170119
        /// author              AGL027
        /// </remarks>
        /// <param name="url">contains pre-prepared url to access the web service</param>
        private abcResponse _downloadStringFromURL(String sUrl)
        {
            abcResponse abcResp = new abcResponse();

            try
            {
                using (WebClient wc = new WebClient())
                {
                    wc.Headers.Add("accept", "application/xml");
                    abcResp.content = wc.DownloadString(sUrl);
                }

            }
            catch (WebException ex)
            {
                var resp = (HttpWebResponse)ex.Response;
                // TODO - 
                abcResp.text = resp.Headers["X-ABC-Error"];
                abcResp.code = resp.Headers["X-ABC-ErrorCode"];
                if (!string.IsNullOrEmpty(abcResp.code))
                {
                    System.Windows.Forms.MessageBox.Show(abcResp.text, "Distance error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
                if (ex.Status == WebExceptionStatus.ProtocolError && ex.Response != null)
                {
                    abcResp.httpcode = "HttpWebResponse: " + (int)resp.StatusCode + ":" + resp.StatusCode.ToString();
                }

            }
            return abcResp;
        }



        /// <summary>
        /// this method transforms the delimited items received from PowerBuilder and transforms into some sort of querystring segment.
        /// If method is to obtain the map the ports are transformed into port names & parm settings linked between the 2 methods are managed here.
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20161205
        /// last modified       20171030
        /// author              AGL027
        /// </remarks>
        /// 
        /// <param name="sDelimitedPorts">a delimited list of port codes from Tramos client</param>
        /// <param name="sDelimitedOpenDefault"></param>
        /// <param name="sDelimitedCloseDefault"></param>
        /// <param name="sDelimitedOpenLegLevel">a delimted list of open ports from the Tramos client.  They may or might not include leg index in format SUZ-0|SUZ-1</param>
        /// <param name="sDelimitedCloseLegLevel">a delimted list of closed ports from the Tramos client.  Optional leg index once again in format DOV-0|DOV-2</param>
        /// <param name="iScanEca">always scan eca zones and provide data in XML</param>
        /// <param name="iEnvNavReg">if method is 'voyage' use envnavreg parm name otherwise use showsecazones.  There is link between the map & voyage option </param>
        /// <param name="iAntiPiracy">if method is 'voyage' use antipiracy parm name otherwise use showpiracyzones.  There is link between the map & voyage option </param>
        public String TransformToUrl(String sDelimitedPorts, String sDelimitedOpenDefault, String sDelimitedCloseDefault, String sDelimitedOpenLegLevel, String sDelimitedCloseLegLevel, int iScanEca, int iEnvNavReg, int iAntiPiracy)
        {
            StringBuilder portParms = new StringBuilder("");
            StringBuilder openLegLevelParms = new StringBuilder("");
            StringBuilder closeLegLevelParms = new StringBuilder("");
            StringBuilder openDefaultParms = new StringBuilder("");
            StringBuilder closeDefaultParms = new StringBuilder("");
            StringBuilder additionalParms = new StringBuilder();
            StringBuilder parturl = new StringBuilder("");

            List<string> portList = sDelimitedPorts.Split('|').ToList();
            List<string> openDefaultList = sDelimitedOpenDefault.Split('|').ToList();
            List<string> closeDefaultList = sDelimitedCloseDefault.Split('|').ToList();
            List<string> openLegLevelList = sDelimitedOpenLegLevel.Split('|').ToList();
            List<string> closeLegLevelList = sDelimitedCloseLegLevel.Split('|').ToList();

            bool firstArg = true;

            /* first the ports that direct each leg inside voyage */
            foreach (var port in portList)
            {
                if (port != "")
                {
                    if (!firstArg)
                    {
                        portParms.Append("&");
                    }
                    else
                    {
                        firstArg = false;
                    }
                    portParms.Append("port=").Append(port);
                }
            }

            /* next the open default routing points */
            foreach (var routingpoint in openDefaultList)
            {
                if (routingpoint != "")
                {
                    openDefaultParms.Append("&open=").Append(routingpoint);
                }
            }
            /* next the close default routing points */
            foreach (var routingpoint in closeDefaultList)
            {
                if (routingpoint != "")
                {
                    closeDefaultParms.Append("&close=").Append(routingpoint);
                }
            }
            
            /* next the open routing points */
            foreach (var routingpoint in openLegLevelList)
            {
                if (routingpoint != "")
                {
                    openLegLevelParms.Append("&open=").Append(routingpoint);
                }
            }
            /* next the close routing points */
            foreach (var routingpoint in closeLegLevelList)
            {
                if (routingpoint != "")
                {
                    closeLegLevelParms.Append("&close=").Append(routingpoint);
                }
            }
            /* last part are the additional parms ('scaneca'; 'envnavreg'; 'antipiracy' */
            if (iScanEca == 1)
            {
                additionalParms.Append("&scaneca=true");
            }
            if (iEnvNavReg == 0)
            {
                additionalParms.Append("&envnavreg=false");
            }
            if (iAntiPiracy == 0)
            {
                additionalParms.Append("&antipiracy=false");
            }
            additionalParms.Append("&illustrate=true&waypointResolution=detailed-high");
            /* concatinate all the parameters needed and pass back to calling process */
            parturl.Append(portParms).Append(openDefaultParms).Append(closeDefaultParms).Append(openLegLevelParms).Append(closeLegLevelParms).Append(additionalParms);
            //System.Windows.Forms.MessageBox.Show(parturl.ToString());
            return parturl.ToString();
        }

        /// <summary>
        /// direct call to return a simple string containing total distance in voyage
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20161205
        /// last modified       20171101
        /// author              AGL027
        /// </remarks>
        /// <param name="sWsParm">the web service querystring parameter from Tramos</param>       
        public String GetDistance(String sWsParm)
        {
            StringBuilder url = new StringBuilder();
            url.Append(_wsns.ToString()).Append("/Distance?").Append(sWsParm).Append("&api_key=" + _apikey);
            abcResponse abcResp = new abcResponse();
            abcResp = _downloadStringFromURL(url.ToString());

            if (!string.IsNullOrEmpty(abcResp.httpcode))
            {
                return abcResp.httpcode;
            }
            return abcResp.content;
        }

        /// <summary>
        /// called to transform the calculations stored state.
        /// the plan is to use just the string and the ports to obtain existing items on a leg level, possibly set these as closed.  
        /// Next obtain the open ports by using the full query. replacing or adding new items to list accordingly
        /// TODO: include also anti-piracy flag
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20170124
        /// last modified       20170407
        /// author              AGL027
        /// </remarks>
        public String GetQueryStringFromAbcEngineState(String sDelimitedPorts, String sAbcEngineState, int iScanEca)
        {
            _routingsForDW = new StringBuilder("");
            List<string> portList = sDelimitedPorts.Split('|').ToList();
            List<string> openRoutings = new List<string>();
            List<string> closedRoutings = new List<string>();
            List<string> finalRoutings = new List<string>();
            String portParms = "";

            String useEcaZone = "&envnavreg=true";
            if (iScanEca == 0)
            {
                useEcaZone = "&envnavreg=false";
            }

            bool firstArg = true;

            foreach (var port in portList)
            {
                if (port != "")
                {
                    if (!firstArg)
                    {
                        portParms += "&";
                    }
                    else
                    {
                        firstArg = false;
                    }
                    portParms += "port=" + port;

                }
            }

            /* get all routings without any parms */
            StringBuilder urlRoutings = new StringBuilder();

            urlRoutings.Append(_wsns.ToString()).Append("/Voyage?").Append(portParms).Append("&illustrate=true&waypointResolution=detailed-high").Append(useEcaZone).Append("&api_key=" + _apikey);
            abcResponse abcResp = new abcResponse();

            abcResp = _downloadStringFromURL(urlRoutings.ToString());
            if (!string.IsNullOrEmpty(abcResp.code))
            {
                System.Windows.Forms.MessageBox.Show(abcResp.code);
                return abcResp.code;
            }
            XDocument closedRoutingsXml = XDocument.Parse(abcResp.content);
            var legs = (from e in closedRoutingsXml.Descendants(_xmlns + "Legs").Elements(_xmlns + "Leg")
                        select new Leg()
                        {
                            WayPointList = e.Element(_xmlns + "Waypoints").Elements(_xmlns + "Waypoint")
                            .Select(r => new WayPoint()
                            {
                                name = r.Element(_xmlns + "Name") != null ? r.Element(_xmlns + "Name").Value : "",
                                routingPointCode = r.Element(_xmlns + "RoutingPoint") != null ? r.Element(_xmlns + "RoutingPoint").Value : "",
                            }).ToList()

                        }).ToList();

            int legIndex = 0;
            foreach (Leg leg in legs)
            {
                foreach (WayPoint wp in leg.WayPointList)
                {
                    if (wp.routingPointCode != "")
                    {
                        closedRoutings.Add(wp.routingPointCode + "-" + legIndex.ToString());
                    }
                }
                legIndex++;
            }

            StringBuilder url = new StringBuilder();
            /* locate the open routing points */
            url.Append(_wsns.ToString()).Append("/Voyage?").Append(portParms).Append("&illustrate=true&waypointResolution=detailed-high").Append("&routingString=").Append(sAbcEngineState).Append(useEcaZone).Append("&api_key=" + _apikey);

            abcResp = _downloadStringFromURL(url.ToString());
            if (!string.IsNullOrEmpty(abcResp.code))
            {
                return abcResp.code;
            }
            _voyagexml = XDocument.Parse(abcResp.content);

            legs = (from e in _voyagexml.Descendants(_xmlns + "Legs").Elements(_xmlns + "Leg")
                    select new Leg()
                    {
                        WayPointList = e.Element(_xmlns + "Waypoints").Elements(_xmlns + "Waypoint")
                        .Select(r => new WayPoint()
                        {
                            name = r.Element(_xmlns + "Name") != null ? r.Element(_xmlns + "Name").Value : "",
                            routingPointCode = r.Element(_xmlns + "RoutingPoint") != null ? r.Element(_xmlns + "RoutingPoint").Value : "",
                        }).ToList()

                    }).ToList();

            legIndex = 0;
            foreach (Leg leg in legs)
            {
                foreach (WayPoint wp in leg.WayPointList)
                {
                    if (wp.routingPointCode != "")
                    {
                        openRoutings.Add(wp.routingPointCode + "-" + legIndex.ToString());
                    }
                }
                legIndex++;
            }

            var differences = closedRoutings.Except(openRoutings);
            foreach (var difference in differences)
            {
                finalRoutings.Add("close=" + difference.ToString());
                String shortCode = difference;

                _routingsForDW.Append(difference.ToString()).Append("\t").Append(GetRPDataByShortCode(shortCode.Substring(0, 3),1)).Append("\t").Append("0").Append("\t").Append(_getOpenByDefault(shortCode.Substring(0, 3))).Append("\r\n");
            }
            
            foreach (String openRouting in openRoutings)
            {
                finalRoutings.Add("open=" + openRouting);
                String shortCode = openRouting;

                _routingsForDW.Append(openRouting.ToString()).Append("\t").Append(GetRPDataByShortCode(shortCode.Substring(0, 3),1)).Append("\t").Append("1").Append("\t").Append(_getOpenByDefault(shortCode.Substring(0, 3))).Append("\r\n");
            }

            string finalRoutingParms = String.Join("&", finalRoutings.ToArray());
            _wsparmstring = portParms + "&" + finalRoutingParms + useEcaZone;
            return _wsparmstring.ToString();
        }

        /// <summary>
        /// returns routings data to be received by datawindow
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20170131
        /// last modified       20170912
        /// author              AGL027
        /// </remarks>
        public String GetRoutingDWFromEngineState()
        {
            return _routingsForDW.ToString();
        }

        /// <summary>
        /// AtoBviaC have a set of default Open and Closed ports.  This method provides a querystring segment that is used on first presentation of routing 
        /// 
        /// date created        20161129
        /// last modified       20161208
        /// author              AGL027
        /// </summary>
        public String GetRPbyType(int iRoutingPointType, String sWsParm)
        {
            StringBuilder openByDefaultFlag = new StringBuilder();

            if (iRoutingPointType == 1) // opened 
            {
                openByDefaultFlag.Append("false");
            }
            else
            {
                openByDefaultFlag.Append("true");
            }

            var rps = from rp in _routingpointxml_cached.Descendants(_xmlns + "RoutingPoint")
                      where rp.Element(_xmlns + "OpenByDefault").Value.Contains(openByDefaultFlag.ToString())
                      select new
                      {
                          RoutingPointCode = rp.Element(_xmlns + "ShortCode").Value,
                      };

            StringBuilder selectedRPs = new StringBuilder();

            foreach (var rp in rps)
            {
                selectedRPs.Append("&").Append(sWsParm).Append("=").Append(rp.RoutingPointCode);
            }
            return selectedRPs.ToString();
        }
        
        /// <summary>
        /// called to update the voyage from the client.
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20161206
        /// last modified       20171101
        /// author              AGL027
        /// </remarks>
        /// <param name="sWsParm"></param>
        /// <param name="iReturnType"></param>
        public String GetVoyage(String sWsParm, int iReturnType)
        {
            String content = "";
            if (sWsParm != _wsparmstring)
            {
                StringBuilder url = new StringBuilder();
                url.Append(_wsns.ToString()).Append("/Voyage?").Append(sWsParm).Append("&illustrate=true&waypointResolution=detailed-high").Append("&api_key=" + _apikey);
                abcResponse abcResp = new abcResponse();
                abcResp = _downloadStringFromURL(url.ToString());
                if (!string.IsNullOrEmpty(abcResp.code))
                {
                    return abcResp.code;
                }
                _wsparmstring = sWsParm;
                this._voyagexml = XDocument.Parse(abcResp.content);
            }
            else
            {
                content = this._voyagexml.ToString();
            }

            if (iReturnType == (int)DataFormat.XML)
            {
                return content;
            }
            else if (iReturnType == (int)DataFormat.PBDataWindowSyntax)
            {
                /* construct data that will be loaded into the calling datawindow */
                StringBuilder dwContent = new StringBuilder();
                dwContent = _parseXMLContentToDWFormat();
                return dwContent.ToString();
            }
            else if (iReturnType == (int)DataFormat.None)
            {
                return "";
            }
            else if (iReturnType == (int)DataFormat.wsParmString)
            {
                return _wsparmstring;
            }
            return "error, unhandled";
        }
        
        /// <summary>
        /// Given a single leg & voyage string (querystring) get the routing points inside.  might be called from within from one of the overridden methods or potentially directly from Tramos
        /// TODO - perhaps work a better means to share the name vars of routing points
        ///      - Possibly better control of the wsparmstring and the options 
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20161209
        /// last modified       20170407
        /// author              AGL027
        /// </remarks>
        public String GetRoutingPointsForSelectedLeg(int iJourneyId, String sWsParm)
        {
            _rpShortCodesInLeg.Clear();
            _rpNamesInLeg.Clear();
            _rpOpenByDefaultInLeg.Clear();

            _routingsForDW = new StringBuilder("");

            if (sWsParm != _wsparmstring)
            {
                String returnCode = this.GetVoyage(sWsParm, 0);
                if (returnCode != "")
                {
                    return returnCode;
                }
            }

            var leg = from voyage in _voyagexml.Descendants(_xmlns + "Leg").Skip(iJourneyId - 1).Take(1)
                      select voyage.Element(_xmlns + "Waypoints");

            StringBuilder routingPointsWithinLeg = new StringBuilder();
            foreach (var item in leg.Elements(_xmlns + "Waypoint"))
            {
                String rpShortCode = item.Element(_xmlns + "RoutingPoint") != null ? item.Element(_xmlns + "RoutingPoint").Value : "";
                if (rpShortCode != "")
                {
                    _routingsForDW.Append(rpShortCode.ToString()).Append("\t").Append(GetRPDataByShortCode(rpShortCode,1)).Append("\t").Append("0").Append("\t").Append(_getOpenByDefault(rpShortCode)).Append("\r\n");
                }
            }
            return _routingsForDW.ToString();
        }



        /// <summary>
        /// Get the full routing point name of Routing Point ShortName that is passed in.
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20161213
        /// last modified       20171023
        /// author              AGL027
        /// </remarks>
        /// 
        /// <param name="sRoutingPointShortCode">AtoBviaC routing port code</param>
        /// <param name="iType">1 = only return the name; 2 = return '|' delimted data</param>
        public String GetRPDataByShortCode(String sRoutingPointShortCode, int iType)
        {
            StringBuilder sData = new StringBuilder("");
            if (iType == 1) { 
                var rpname = (from rp in _routingpointxml_cached.Descendants(_xmlns + "RoutingPoint")
                          where rp.Element(_xmlns + "ShortCode").Value.Contains(sRoutingPointShortCode)
                          select rp.Element(_xmlns + "Name").Value).FirstOrDefault();

                sData.Append(rpname.ToString());

            } else
            {
                var rpdata = from rp in _routingpointxml_cached.Descendants(_xmlns + "RoutingPoint")
                          where rp.Element(_xmlns + "ShortCode").Value.Contains(sRoutingPointShortCode)
                          select new ActiveRoutingPoint
                          {
                              ShortCode = sRoutingPointShortCode, 
                              Name = rp.Element(_xmlns + "Name").Value,
                              IsOpen = Convert.ToBoolean(rp.Element(_xmlns + "OpenByDefault").Value),
                          };
                
               foreach(ActiveRoutingPoint item in rpdata)
               {
                    sData.Append(item.ShortCode).Append("|").Append(item.Name).Append("|");
                    if (item.IsOpen.ToString().ToLower() == "true")
                    {
                        sData.Append("1");
                    } else {
                        sData.Append("0");
                    }
                    
               }
            }
            return sData.ToString();
        }

        /// <summary>
        /// This method is designed to extract the routing data found within the querystring
        /// called from only inside the calculation module
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20170206
        /// last modified       20170912
        /// author              AGL027
        /// </remarks>
        public string GetRoutingDWfromWSParm()
        {
            _routingsForDW.Clear();

            NameValueCollection parmData = HttpUtility.ParseQueryString(this._wsparmstring);
            var items = parmData.AllKeys.SelectMany(parmData.GetValues, (k, v) => new { key = k, value = v });

            foreach (var item in items)
            {
                if (item.key == "open")
                {
                    _routingsForDW.Append(item.value).Append("\t").Append(GetRPDataByShortCode(item.value.Substring(0, 3),1)).Append("\t").Append("1").Append("\t").Append(_getOpenByDefault(item.value.Substring(0, 3))).Append("\r\n");
                }
                else if (item.key == "close")
                {
                    _routingsForDW.Append(item.value).Append("\t").Append(GetRPDataByShortCode(item.value.Substring(0, 3),1)).Append("\t").Append("0").Append("\t").Append(_getOpenByDefault(item.value.Substring(0, 3))).Append("\r\n");
                }
            }
            return _routingsForDW.ToString();
        }

        /// <summary>
        /// This method with public scope is called by consumer to obtain data in GEOJSON format for the
        /// Leaflet.API - that in turn generates the interactive map.   
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20171019
        /// last modified       20171031
        /// author              AGL027
        /// </remarks>
        public string GetGeoJsonData(string sWsParm)
        {
            StringBuilder sGeoJsonData = new StringBuilder("");
            StringBuilder sAbcLegRoutingFeaturePre = new StringBuilder("");
            StringBuilder sAbcMarkerRPs = new StringBuilder("");

            double dPreviousLon = 0;

            OffSetRouting[] offsets =
            {
                new OffSetRouting { OffSetAmount = -360, JSVariableName="abc_routing_neg360"  },
                new OffSetRouting { OffSetAmount = 0, JSVariableName="abc_routing"  },
                new OffSetRouting { OffSetAmount = 360, JSVariableName="abc_routing_pos360"  }
            };

            OffSetPortMarkers[] markerOffsets =
            {
                new OffSetPortMarkers { OffSetAmount = -360, JSVariableName="abc_marker_ports_neg360"  },
                new OffSetPortMarkers { OffSetAmount = 0, JSVariableName="abc_marker_ports"  },
                new OffSetPortMarkers { OffSetAmount = 360, JSVariableName="abc_marker_ports_pos360"  }
            };

            var legs = (from e in _voyagexml.Descendants(_xmlns + "Legs").Elements(_xmlns + "Leg")
                        select new Leg()
                        {

                            fromPort = e.Elements(_xmlns + "FromPort")
                            .Select(r => new Port()
                            {
                                name = (string)r.Element(_xmlns + "Name"),
                                code = (string)r.Element(_xmlns + "Code"),
                                LatGeodetic = double.Parse(r.Element(_xmlns + "LatGeodetic").Value, CultureInfo.InvariantCulture),
                                Lon = double.Parse(r.Element(_xmlns + "Lon").Value, CultureInfo.InvariantCulture)
                            }).FirstOrDefault(),

                            toPort = e.Elements(_xmlns + "ToPort")
                            .Select(r => new Port()
                            {
                                name = r.Element(_xmlns + "Name") != null ? r.Element(_xmlns + "Name").Value : "",
                                code = r.Element(_xmlns + "Code") != null ? r.Element(_xmlns + "Code").Value : "",
                                LatGeodetic = double.Parse(r.Element(_xmlns + "LatGeodetic").Value, CultureInfo.InvariantCulture),
                                Lon = double.Parse(r.Element(_xmlns + "Lon").Value, CultureInfo.InvariantCulture)

                            }).First(),

                            WayPointList = e.Element(_xmlns + "Waypoints").Elements(_xmlns + "Waypoint")
                            .Select(r => new WayPoint()
                            {
                                name = r.Element(_xmlns + "Name") != null ? r.Element(_xmlns + "Name").Value : "",
                                DistanceFromStart = Convert.ToDecimal(r.Element(_xmlns + "DistanceFromStart").Value),
                                routingPointCode = r.Element(_xmlns + "RoutingPoint") != null ? r.Element(_xmlns + "RoutingPoint").Value : "",
                                EcaZoneToPrevious = r.Element(_xmlns + "EcaZoneToPrevious") != null ? r.Element(_xmlns + "EcaZoneToPrevious").Value : "",
                                LatGeodetic = double.Parse(r.Element(_xmlns + "LatGeodetic").Value, CultureInfo.InvariantCulture),
                                Lon = double.Parse(r.Element(_xmlns + "Lon").Value, CultureInfo.InvariantCulture)

                            }).ToList(),

                            PolyLineList = e.Element(_xmlns + "IllustrativeRoute").Elements(_xmlns + "Polyline").Elements(_xmlns + "Point")
                            .Select(r => new PolyLinePoint()
                            {
                                LatGeodetic = double.Parse(r.Element(_xmlns + "LatGeodetic").Value, CultureInfo.InvariantCulture),
                                Lon = double.Parse(r.Element(_xmlns + "Lon").Value, CultureInfo.InvariantCulture)

                            }).ToList()

                        }).ToList();

            foreach (OffSetRouting o in offsets)
            {
                o.AbcRouting.Append(@"
  var ").Append(o.JSVariableName).Append(@" = {
    ""type"": ""FeatureCollection"",
    ""features"": [
");
            }

            /* on a port level we add the detail to this collection */
            foreach (OffSetPortMarkers o in markerOffsets)
            {
                o.AbcMarkerPorts.Append(@"
var ").Append(o.JSVariableName).Append(@" = {
  ""type"": ""FeatureCollection"",
  ""features"": [
            ");
}


            sAbcMarkerRPs.Append(@"
var abc_routing_points={
  ""type"": ""FeatureCollection"",
  ""features"": [
");

            sAbcLegRoutingFeaturePre.Append(@"
{
    ""type"": ""Feature"",
        ""geometry"": {
        ""type"": ""LineString"",
        ""coordinates"": [");

            int iLegIndex = 0;
            int iLegCount = legs.Count();
            decimal dDistanceFromStart = 0;

            double dAdjustment = 0;

            List<ActiveRoutingPoint> ActiveRPsList = new List<ActiveRoutingPoint>();
            
            /* now the model has been populated we try and construct the route */
            foreach (Leg leg in legs)
            {
                List<ActiveRoutingPoint> RPsInLeg = new List<ActiveRoutingPoint>();
                bool bFirstPoint = true;

                if (leg.fromPort.code != leg.toPort.code)
                {
                    RPsInLeg = GetRPList(iLegIndex, sWsParm);
                }
                string sFromCode = leg.fromPort.code;
                string sToCode = leg.toPort.code;
                StringBuilder sAbcLegRoutingFeaturePost = new StringBuilder();
                /* use PolyLinePoint's */
                foreach (PolyLinePoint p in leg.PolyLineList)
                {
                    if (bFirstPoint)
                    {
                        dPreviousLon = p.Lon;
                        bFirstPoint = false;
                    }
                    if (p.Lon > (dPreviousLon + 180))
                    {
                        dAdjustment -= 360;
                    }
                    else if (p.Lon < (dPreviousLon - 180))
                    {
                        dAdjustment += 360;
                    }
                    dPreviousLon = p.Lon;
                    p.Lon += dAdjustment;
                    foreach (OffSetRouting o in offsets)
                    {
                        o.CoordContent.Append("[").Append((p.Lon + o.OffSetAmount).ToString()).Append(",").Append(p.LatGeodetic.ToString()).Append("],");
                    }
                }

                /* Next loop through Way Points */
                foreach (WayPoint wp in leg.WayPointList)
                {
                    bool bShowPortPopup = false;
                    if (wp.name == sFromCode)
                    {
                        wp.name = leg.fromPort.name;
                        if (iLegIndex == 0)
                        {
                            bShowPortPopup = true;
                        }
                    }
                    if (wp.name == sToCode || wp.name == leg.toPort.name)
                    {
                        wp.name = leg.toPort.name;
                        bShowPortPopup = true;
                        sAbcLegRoutingFeaturePost.Append(@"
]
    },
    ""properties"": {
        ""popupContent"": ""<h2>").Append(@wp.name).Append(@"<h2>"",
        ""Description"": """).Append(leg.fromPort.name).Append(@" to ").Append(leg.toPort.name).Append(@"""
},
""id"": ").Append(iLegIndex.ToString()).Append(@"
},
");
                    }

                    if (wp.routingPointCode != "")
                    { 
                        foreach (ActiveRoutingPoint rp in RPsInLeg.Where(x => x.ShortCode.Contains(wp.routingPointCode)))
                        {
                            rp.IsOpen = true;
                            rp.IsAdvanced = false;
                            /* adjust the coordinates so routing point is found on routing line */
                            rp.LatGeodetic = wp.LatGeodetic;
                            rp.Lon = wp.Lon;
                        }
                    }
                    if (bShowPortPopup)
                    {
                        dDistanceFromStart += wp.DistanceFromStart;
                        foreach (OffSetPortMarkers o in markerOffsets)
                        {
                            o.AbcMarkerPorts.Append(@"
                            {
                               ""type"": ""Feature"",
                                  ""geometry"": {
                                    ""type"": ""Point"",
                                    ""coordinates"": [").Append(wp.Lon + o.OffSetAmount).Append(",").Append(wp.LatGeodetic).Append(@"
                                    ]
                                },
                                  ""properties"": {
                                    ""PortName"": """).Append(wp.name).Append(@""",
                                    ""Lat"": ").Append(String.Format("{0:0.00}", wp.LatGeodetic)).Append(@",
                                    ""Lon"": ").Append(String.Format("{0:0.00}", wp.Lon)).Append(@",
                                    ""DistanceFromStart"": ").Append(String.Format("{0:0.00}", dDistanceFromStart)).Append(@",
                                    ""LegDistance"": ").Append(String.Format("{0:0.00}", wp.DistanceFromStart)).Append(@"
                                  }
                                },
                            ");
                        }
                        
                    }
                }
                foreach (OffSetRouting o in offsets)
                {
                    o.AbcRouting.Append(sAbcLegRoutingFeaturePre).Append(o.CoordContent).Append(sAbcLegRoutingFeaturePost);
                    o.CoordContent = new StringBuilder("");
                }
                ActiveRPsList.AddRange(RPsInLeg);
                iLegIndex++;
            }
            
            /* After collecting all the routing points we can now format for each leg */
            foreach (ActiveRoutingPoint rp in ActiveRPsList)
            {
                sAbcMarkerRPs.Append(@"
                {
                    ""type"": ""Feature"",
                        ""geometry"": {
                        ""type"": ""Point"",
                        ""coordinates"": [").Append(rp.Lon.ToString()).Append(",").Append(rp.LatGeodetic.ToString()).Append(@"
                        ]
                    },
                        ""properties"": {
                        ""ShortCode"": """).Append(rp.ShortCode).Append(@""",
                        ""Name"": """).Append(rp.Name).Append(@""",
                        ""LegIndex"": ").Append(rp.LegIndex.ToString()).Append(@",
                        ""Open"": ").Append(rp.IsOpen.ToString().ToLower()).Append(@",
                        ""Advanced"": ").Append(rp.IsAdvanced.ToString().ToLower()).Append(@"
                        }
                    },
                ");
            }

            /* Close each JS struture */
            sAbcMarkerRPs.Append(@"
    ]
};
");

            foreach (OffSetRouting o in offsets)
            {
                o.AbcRouting.Append(@"
    ]
};
");
            }

            foreach (OffSetPortMarkers o in markerOffsets)
            {
                o.AbcMarkerPorts.Append(@"
	]	
};
");
            }

            foreach (OffSetRouting o in offsets)
            {
                sGeoJsonData.Append(@o.AbcRouting);
            }
            foreach (OffSetPortMarkers o in markerOffsets)
            {
                sGeoJsonData.Append(@o.AbcMarkerPorts);
            }
            sGeoJsonData.Append(sAbcMarkerRPs);
            return sGeoJsonData.ToString();
        }

        /// <summary>
        /// This method creates a list per Leg that is called.  It loads the AtoBviaC defaults first and then
        /// Overrides those with whatever Tramos has.
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20171023
        /// last modified       20171027
        /// author              AGL027
        /// </remarks>
        public List<ActiveRoutingPoint> GetRPList(int iLegIndex, string sWsParm)
        {
            List<ActiveRoutingPoint> rplisting = new List<ActiveRoutingPoint>();

            /* Phase I - We get AtoBviaC defaults and initialize the list */
            var rps = from rp in _routingpointxml_cached.Descendants(_xmlns + "RoutingPoint")
                     select new ActiveRoutingPoint
                     {
                         Type = "Point",
                         ShortCode = rp.Element(_xmlns + "ShortCode").Value,
                         LatGeodetic = double.Parse(rp.Element(_xmlns + "LatGeodetic").Value),
                         Lon = double.Parse(rp.Element(_xmlns + "Lon").Value),
                         Name = rp.Element(_xmlns + "Name").Value,
                         IsOpen = Convert.ToBoolean(rp.Element(_xmlns + "OpenByDefault").Value),
                         IsAdvanced = true,
                         LegIndex = iLegIndex
                     };

           rplisting.AddRange(rps);

           /* Phase II - with the wsparm received from Tramos we load what is set there in terms of PCGROUP override and what we have already */
            NameValueCollection parmData = HttpUtility.ParseQueryString(sWsParm);
            var items = parmData.AllKeys.SelectMany(parmData.GetValues, (k, v) => new { key = k, value = v });
            /* important that voyage wide routing point control is passed before leg level detail.  This > open=DOV&close=DOV-0, instead of close=DOV-0&open=DOV */
            var sorteditems = items.OrderBy(kvp => kvp.value);

            foreach (var item in sorteditems)
            {
                bool bIsOpen = true;
                bool bIsAdvanced = false;
                if (item.key == "close")
                {
                    bIsOpen = false;
                    bIsAdvanced = true;
                }
                if (item.value.Length == 3 && (item.key=="close" || item.key == "open")) 
                {
                    foreach (ActiveRoutingPoint rp in rplisting.Where(x => x.ShortCode.Contains(item.value)))
                    {
                        rp.IsOpen = bIsOpen;
                        rp.IsAdvanced = bIsAdvanced;
                    }
                }
                else if (item.value.Length > 3 && (item.key == "close" || item.key == "open"))
                {
                    int iLegindex = Convert.ToInt16(item.value.Substring(4));
                    foreach (ActiveRoutingPoint rp in rplisting.Where(x => x.ShortCode.Contains(item.value.Substring(0,3)) && x.LegIndex == iLegindex))
                    {
                        rp.IsOpen = bIsOpen;
                        rp.IsAdvanced = bIsAdvanced;
                    }
                }
            }
            return rplisting;
        }
    }
}
