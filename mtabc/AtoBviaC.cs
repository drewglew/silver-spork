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

/* CR4343 */
namespace tankers.distances
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("mt.distances.AtoBviaC")]

    /* created 20161205 */
    public class AtoBviaC
    {
        private static XNamespace _xmlns;
        private static XNamespace _wsns;
        private String _apikey;
        private String _wsparmstring;
        private XDocument _routingpointxml_cached;
        private XDocument _voyagexml;
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
        /// this must be called by the requestor to load the necessary api_key that is needed by AtoBviaC.
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20161215
        /// last modified       20171215
        /// author              AGL027
        /// </remarks>
        /// 
        /// <param name="apikey">api_key must be stored in client application.</param> 
        public void SetApiKey(String apikey)
        {
            _apikey = apikey;
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
        /// 
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
        /// 
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
                if (rp.openbydefault.ToUpper()=="TRUE")
                {
                    routingList.Append("1").Append("\r\n");
                } else
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
        /// 
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

            //System.Windows.Forms.MessageBox.Show(portList.ToString());

            return portList.ToString();
        }



        /// <summary>
        /// low level communication to web service; this is to handle image processing
        /// </summary>
        /// 
        /// <remarks>
        /// date created        20161205
        /// last modified       20161205
        /// author              AGL027
        /// </remarks>
        /// 
        /// <param name="errorMessage">return the error message if try catches error</param>
        /// <param name="url">contains pre-prepared url to access the web service</param>
        private byte[] _downloadDataFromURL(String url, ref String errorMessage)
        {
            byte[] content = new byte[0];

            try
            {
                using (WebClient wc = new WebClient())
                    content = wc.DownloadData(url);

            }
            catch (WebException ex)
            {
                var resp = (HttpWebResponse)ex.Response;

                string errorHeaderText = resp.Headers["X-ABC-Error"];
                string errorHeaderCode = resp.Headers["X-ABC-ErrorCode"];

                System.Windows.Forms.MessageBox.Show(errorHeaderText + "|" + errorHeaderCode);

                if (errorHeaderText != null && errorHeaderCode != null)
                {
                    errorMessage = errorHeaderCode + "|" + errorHeaderText;
                }
                else
                {
                    if (ex.Status == WebExceptionStatus.ProtocolError && ex.Response != null)
                    {
                        errorMessage = "HttpWebResponse: " + (int)resp.StatusCode + ":" + resp.StatusCode.ToString();
                        System.Windows.Forms.MessageBox.Show(errorMessage.ToString());
                    }
                }

            }
            return content;
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
        /// <param name="errorMessage">contains pre-prepared url to access the web service</param>
        private abcResponse _downloadStringFromURL(String url)
        {
            abcResponse abcResp = new abcResponse();

            try
            {
                using (WebClient wc = new WebClient())
                {
                    wc.Headers.Add("accept", "application/xml");
                    abcResp.content = wc.DownloadString(url);
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
        /// TODO - expand to handle partial routing points?
        /// 
        /// date created        20161205
        /// last modified       20171016
        /// author              AGL027
        /// </remarks>
        /// <param name="method">possibly could be Image; Voyage</param>
        /// <param name="ports">a delimited list of port codes from Tramos client</param>
        /// <param name="open_default"></param>
        /// <param name="close_default"></param>
        /// <param name="open_leglevel">a delimted list of open ports from the Tramos client.  They may or might not include leg index in format SUZ-0|SUZ-1</param>
        /// <param name="close_leglevel">a delimted list of closed ports from the Tramos client.  Optional leg index once again in format DOV-0|DOV-2</param>
        /// <param name="scaneca">always scan eca zones and provide data in XML</param>
        /// <param name="envnavreg">if method is 'voyage' use envnavreg parm name otherwise use showsecazones.  There is link between the map & voyage option </param>
        /// <param name="antipiracy">if method is 'voyage' use antipiracy parm name otherwise use showpiracyzones.  There is link between the map & voyage option </param>
        public String TransformToUrl(String method, String ports, String open_default, String close_default, String open_leglevel, String close_leglevel, int scaneca, int envnavreg, int antipiracy)
        {
            String portArg = "";

            StringBuilder portParms = new StringBuilder("");
            StringBuilder openParms = new StringBuilder("");
            StringBuilder closeParms = new StringBuilder("");
            StringBuilder openDefaultParms = new StringBuilder("");
            StringBuilder closeDefaultParms = new StringBuilder("");
            StringBuilder additionalParms = new StringBuilder();
            StringBuilder parturl = new StringBuilder("");

            List<string> portList;
            List<string> openDefaultList;
            List<string> closeDefaultList;
            List<string> openList;
            List<string> closeList;

            portList = ports.Split('|').ToList();
            openDefaultList = open_default.Split('|').ToList();
            closeDefaultList = close_default.Split('|').ToList();
            openList = open_leglevel.Split('|').ToList();
            closeList = close_leglevel.Split('|').ToList();

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

                    if (method.ToLower() == "image")
                    {
                        if (_portsxml_cached == null)
                        {
                            _portsxml_cached = _getPorts();
                        }
                        // obtain nice port name and escape problems chars i.e. <space> 
                        portArg = Uri.EscapeDataString(_getPortName(port));
                    }
                    else
                    {
                        portArg = port;
                    }
                    portParms.Append("port=").Append(portArg);
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
            foreach (var routingpoint in openList)
            {
                if (routingpoint != "")
                {
                    openParms.Append("&open=").Append(routingpoint);
                }
            }
            /* next the close routing points */
            foreach (var routingpoint in closeList)
            {
                if (routingpoint != "")
                {
                    closeParms.Append("&close=").Append(routingpoint);
                }
            }
            /* last part are the additional parms ('scaneca'; 'envnavreg'; 'antipiracy' */
            if (scaneca == 1)
            {
                additionalParms.Append("&scaneca=true");
            }
            if (envnavreg == 0)
            {
                if (method == "voyage")
                {
                    additionalParms.Append("&envnavreg=false");
                } else if (method == "image")
                {
                    additionalParms.Append("&showsecazones=false&envnavreg=false");
                }
            }
            if (antipiracy == 0)
            {
                if (method == "voyage")
                {
                    additionalParms.Append("&antipiracy=false");
                } else if (method == "image")
                {
                    additionalParms.Append("&showpiracyzones=false&antipiracy=false");
                }
            }

            /* new item since 20171016 */
            additionalParms.Append("&waypointResolution=detailed-high");


            /* concatinate all the parameters needed and pass back to calling process */
            parturl.Append(portParms).Append(openDefaultParms).Append(closeDefaultParms).Append(openParms).Append(closeParms).Append(additionalParms);
            //System.Windows.Forms.MessageBox.Show(parturl.ToString());

            return parturl.ToString();
        }

        /// <summary>
        /// direct call to return a simple string containing total distance in voyage
        /// 
        /// date created        20161205
        /// last modified       20161205
        /// author              AGL027
        /// </summary>
        public String GetDistance(String wsParmString)
        {

            StringBuilder url = new StringBuilder();
            // might be called from within from one of the overridden methods or potentially directly from Tramos
            //_wsparmstring = wsParmString;

            url.Append(_wsns.ToString()).Append("/Distance?").Append(wsParmString).Append("&api_key=" + _apikey);

            abcResponse abcResp = new abcResponse();
            abcResp = _downloadStringFromURL(url.ToString());

            if (!string.IsNullOrEmpty(abcResp.httpcode))
            {
                return abcResp.httpcode;
            }
            else
            {
                // _wsparmstring = _createNewQryStr(content);
            }
            return abcResp.content;
        }


        /// <summary>
        /// direct call to return a simple string containing total distance in voyage
        /// 
        /// date created        20161205
        /// last modified       20170912
        /// author              AGL027
        /// </summary>
        public String GetPortToPortDistance(String wsParmString, int journeyId)
        {
            StringBuilder url = new StringBuilder();
            if (wsParmString != _wsparmstring)
            {
                _wsparmstring = wsParmString;
                url.Append(_wsns.ToString()).Append("/Voyage?").Append(wsParmString).Append("&api_key=" + _apikey);
                abcResponse abcResp = new abcResponse();
                abcResp = _downloadStringFromURL(url.ToString());
                if (!string.IsNullOrEmpty(abcResp.httpcode))
                {
                    return "error";
                }
                _voyagexml = XDocument.Parse(abcResp.content);
            }

            var distance = from e in _voyagexml.Descendants(_xmlns + "Leg").Skip(journeyId - 1).Take(1)
                      select e.Element(_xmlns + "Distance").Value;
            return Convert.ToString(distance);

        }

        /// <summary>
        /// called to transform the calculations stored state.
        /// 
        /// TODO: include also anti-piracy flag
        /// 
        /// date created        20170124
        /// last modified       20170407
        /// author              AGL027
        /// </summary>
        public String GetQueryStringFromAbcEngineState(String ports, String abcEngineState, int scanEca)
        {
            // the plan is to use just the string and the ports to obtain existing items on a leg level, possibly set these as closed.
            // then obtain the open ports by using the full query. replacing or adding new items to list accordingly

            _routingsForDW = new StringBuilder("");

            List<string> portList;
            List<string> openRoutings = new List<string>();
            List<string> closedRoutings = new List<string>();
            List<string> finalRoutings = new List<string>();
            String portParms = "";

            String useEcaZone = "&envnavreg=true";
            if (scanEca == 0) {
                useEcaZone = "&envnavreg=false";
            }

            portList = ports.Split('|').ToList();

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

            /* so we get all routings without any parms */
            StringBuilder urlRoutings = new StringBuilder();
            
            urlRoutings.Append(_wsns.ToString()).Append("/Voyage?").Append(portParms).Append("&waypointResolution=detailed-high").Append(useEcaZone).Append("&api_key=" + _apikey);
            abcResponse abcResp = new abcResponse();

            abcResp = _downloadStringFromURL(urlRoutings.ToString());
            if (!string.IsNullOrEmpty(abcResp.code))
            {
                // HERE
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
            //Now we locate the open routing points
            url.Append(_wsns.ToString()).Append("/Voyage?").Append(portParms).Append("&waypointResolution=detailed-high").Append("&routingString=").Append(abcEngineState).Append(useEcaZone).Append("&api_key=" + _apikey);

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

                _routingsForDW.Append(difference.ToString()).Append("\t").Append(_getRPName(shortCode.Substring(0, 3))).Append("\t").Append("0").Append("\t").Append(_getOpenByDefault(shortCode.Substring(0, 3))).Append("\r\n");
            }


            foreach (String openRouting in openRoutings)
            {
                finalRoutings.Add("open=" + openRouting);
                String shortCode = openRouting;

                _routingsForDW.Append(openRouting.ToString()).Append("\t").Append(_getRPName(shortCode.Substring(0, 3))).Append("\t").Append("1").Append("\t").Append(_getOpenByDefault(shortCode.Substring(0, 3))).Append("\r\n");
            }

            string finalRoutingParms = String.Join("&", finalRoutings.ToArray());

            _wsparmstring = portParms + "&" + finalRoutingParms + useEcaZone;

            //System.Windows.Forms.MessageBox.Show(_routingsForDW.ToString());

            return _wsparmstring.ToString();
        }

        /// <summary>
        /// returns routings data to be received by datawindow
        /// 
        /// date created        20170131
        /// last modified       20170912
        /// author              AGL027
        /// </summary>
        public String GetRoutingDWFromEngineState()
        {
            return _routingsForDW.ToString();
        }
        

        /// <summary>
        /// called to update the voyage from the client.
        /// 
        /// date created        20161206
        /// last modified       20170407
        /// author              AGL027
        /// </summary>
        public String GetVoyage(String wsParmString, int returnType)
        {
            // might be called from within from one of the overridden methods or potentially directly from Tramos
            
            String content = "";

            if (wsParmString != _wsparmstring)
            {
                StringBuilder url = new StringBuilder();
               
                url.Append(_wsns.ToString()).Append("/Voyage?").Append("&waypointResolution=detailed-high").Append(wsParmString).Append("&api_key=" + _apikey);
                abcResponse abcResp = new abcResponse();
                abcResp = _downloadStringFromURL(url.ToString());
                if (!string.IsNullOrEmpty(abcResp.code))
                {
                    return abcResp.code;
                }
                _wsparmstring = wsParmString;
                this._voyagexml = XDocument.Parse(abcResp.content);
            }
            else
            {
                content = this._voyagexml.ToString();
            }

            if (returnType == (int)DataFormat.XML)
            {
                return content;
            }
            else if (returnType == (int)DataFormat.PBDataWindowSyntax)
            {

                /* construct data that will be loaded into the calling datawindow */
                StringBuilder dwContent = new StringBuilder();

                dwContent = _parseXMLContentToDWFormat(content);

                return dwContent.ToString();

            }
            else if (returnType == (int)DataFormat.None)
            {
                return "";
            }
            else if (returnType == (int)DataFormat.wsParmString)
            {
                return _wsparmstring;
            }
            return "error, unhandled";
        }

        /// <summary>
        /// Obtain delimited routing data from the voyage XML string
        /// 
        /// date created        20170116
        /// last modified       20170118
        /// author              AGL027
        /// </summary>
        private StringBuilder _parseXMLContentToDWFormat(String voyageXML)
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

                /* now the fun begins */
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
                        // we are in the first port of the first leg - write out the initial port
                        fromPortData.Append(leg.fromPort.code).Append("\t").Append(leg.fromPort.name).Append("\t").Append("0.000").Append("\t").Append("0.000").Append("\t").Append(is_eca).Append("\r\n");
                    }
                    else if (leg.toPort.code == wp.name)
                    {
                        // we are in the last port of the leg.  We always write this data out.
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
                            wayPointData.Append(wp.routingPointCode).Append("\t").Append(_getRPName(wp.routingPointCode)).Append("\t").Append(sumOfDistancesSinceLastKnown.ToString()).Append("\t").Append(sumOfEcaZoneDistancesSinceLastKnown.ToString()).Append("\t").Append(is_eca).Append("\r\n");
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
        /// Given a single leg & voyage string (querystring) get the routing points inside.  might be called from within from one of the overridden methods or potentially directly from Tramos
        /// TODO - perhaps work a better means to share the name vars of routing points
        /// 
        /// TODO Possibly better control of the wsparmstring and the options 
        /// 
        /// date created        20161209
        /// last modified       20170407
        /// author              AGL027
        /// </summary>
        public String GetRoutingPointsForSelectedLeg(int journeyId, String wsParmString)
        {
            _rpShortCodesInLeg.Clear();
            _rpNamesInLeg.Clear();
            _rpOpenByDefaultInLeg.Clear();

            _routingsForDW = new StringBuilder("");

            if (wsParmString != _wsparmstring)
            {
  
                String returnCode = this.GetVoyage(wsParmString, 0);
                
                if (returnCode != "")
                {
                    return returnCode;
                }
            }

            var leg = from voyage in _voyagexml.Descendants(_xmlns + "Leg").Skip(journeyId - 1).Take(1)
                      select voyage.Element(_xmlns + "Waypoints");

            StringBuilder routingPointsWithinLeg = new StringBuilder();

            foreach (var item in leg.Elements(_xmlns + "Waypoint"))
            {
                String rpShortCode = item.Element(_xmlns + "RoutingPoint") != null ? item.Element(_xmlns + "RoutingPoint").Value : "";

                if (rpShortCode != "")
                {
                    _routingsForDW.Append(rpShortCode.ToString()).Append("\t").Append(_getRPName(rpShortCode)).Append("\t").Append("0").Append("\t").Append(_getOpenByDefault(rpShortCode)).Append("\r\n");
                }

            }
            return _routingsForDW.ToString();
        }

        /// <summary>
        /// Obtain single delimited collection of routing point names inside a string
        /// 
        /// OBSOLETE
        /// 
        /// date created        ??????
        /// last modified       ??????
        /// author              AGL027
        /// </summary>
        public String GetRPNamesInsideLeg()
        {
            return _rpNamesInLeg.ToString();
        }

        /// <summary>
        /// Obtain single delimited collection of routing point short codes inside a string
        /// 
        /// date created        ??????
        /// last modified       ??????
        /// author              AGL027
        /// </summary>
        public String GetRPShortCodesInsideLeg()
        {
            return _rpShortCodesInLeg.ToString();
        }

        /// <summary>
        /// Obtain single delimited collection of routing point short codes that are open inside a string
        /// 
        /// OBSOLETE
        /// 
        /// date created        ??????
        /// last modified       ??????
        /// author              AGL027
        /// </summary>
        public String GetRPOpenByDefaultInsideLeg()
        {
            return _rpOpenByDefaultInLeg.ToString();
        }

        /// <summary>
        /// Direct call to web service that gets the data for all Routing Points and allows us to load this into an
        /// XMLDocument data object.  This method is not dependent on a valid api_key.
        /// TODO - could be optimized
        /// 
        /// date created        20161206
        /// last modified       20170119
        /// author              AGL027
        /// </summary>
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
        /// 
        /// date created        20170119
        /// last modified       20170119
        /// author              AGL027
        /// </summary>
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
        /// Get the full routing point name of Routing Point ShortName that is passed in.
        /// 
        /// date created        20161213
        /// last modified       20161213
        /// author              AGL027
        /// </summary>
        private String _getRPName(String rpShortCode)
        {
            var rpname = (from rp in _routingpointxml_cached.Descendants(_xmlns + "RoutingPoint")
                          where rp.Element(_xmlns + "ShortCode").Value.Contains(rpShortCode)
                          select rp.Element(_xmlns + "Name").Value).FirstOrDefault();

            return rpname.ToString();
        }

        /// <summary>
        /// Get the full port name of Port Code that is passed in.
        /// 
        /// date created        20170119
        /// last modified       20170119
        /// author              AGL027
        /// </summary>
        private String _getPortName(String portCode)
        {
            var portname = (from p in _portsxml_cached.Descendants(_xmlns + "Port")
                            where p.Element(_xmlns + "Code").Value.Contains(portCode)
                            select p.Element(_xmlns + "Name").Value).FirstOrDefault();
            return portname;
        }

        /// <summary>
        /// Simply get OpenByDefault value from the routing point XMLDocument object.
        /// 
        /// date created        ??????
        /// last modified       ??????
        /// author              AGL027
        /// </summary>
        private String _getOpenByDefault(String rpShortCode)
        {
            var rps = (from rp in _routingpointxml_cached.Descendants(_xmlns + "RoutingPoint")
                       where rp.Element(_xmlns + "ShortCode").Value.Contains(rpShortCode)
                       select rp.Element(_xmlns + "OpenByDefault").Value).FirstOrDefault();
            return rps;
        }

        /// <summary>
        /// Simply get open/closed parameters from passed in querystring.
        /// 
        /// date created        20170131
        /// last modified       20170131
        /// author              AGL027
        /// </summary>
        public string GetRPfromQueryStringInLeg(int journeyId, String rpType)
        {

            var parmData = HttpUtility.ParseQueryString(this._wsparmstring).Get(rpType.ToLower());
            StringBuilder returnVal = new StringBuilder("");

            _rpShortCodesInLeg.Clear();
            _rpNamesInLeg.Clear();

            foreach (var parmItem in parmData)
            {
                _rpShortCodesInLeg.Append(parmItem.ToString() + "|");
                _rpNamesInLeg.Append(_getRPName(parmItem.ToString()) + "|");
            }

            return _rpShortCodesInLeg.ToString();
        }

        /// <summary>
        /// This method is designed to extract the routing data found within the querystring
        /// 
        /// date created        20170206
        /// last modified       20170912
        /// author              AGL027
        /// </summary>
        public string GetRoutingDWfromWSParm()
        {
            _routingsForDW.Clear();
            
            NameValueCollection parmData = HttpUtility.ParseQueryString(this._wsparmstring);
            var items = parmData.AllKeys.SelectMany(parmData.GetValues, (k, v) => new { key = k, value = v });

            foreach (var item in items)
            {
                if (item.key == "open")
                {
                    _routingsForDW.Append(item.value).Append("\t").Append(_getRPName(item.value.Substring(0, 3))).Append("\t").Append("1").Append("\t").Append(_getOpenByDefault(item.value.Substring(0, 3))).Append("\r\n");
                }
                else if (item.key == "close")
                {
                    _routingsForDW.Append(item.value).Append("\t").Append(_getRPName(item.value.Substring(0, 3))).Append("\t").Append("0").Append("\t").Append(_getOpenByDefault(item.value.Substring(0, 3))).Append("\r\n");
                }
            }
            return _routingsForDW.ToString();
        }

        /// <summary>
        /// AtoBviaC have a set of default Open and Closed ports.  This method provides a querystring segment that is used on first presentation of routing 
        /// 
        /// date created        20161129
        /// last modified       20161208
        /// author              AGL027
        /// </summary>
        public String GetRPbyType(int rpType, String queryParm)
        {
            StringBuilder openByDefaultFlag = new StringBuilder();

            if (rpType == 1) // opened 
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
                selectedRPs.Append("&").Append(queryParm).Append("=").Append(rp.RoutingPointCode);
            }
            return selectedRPs.ToString();
        }

        /// <summary>
        /// This method with public scope is called by consumer to obtain the map image.  It requires assistance from both  _composeMapOptionString() and _downloadDataFromURL()
        /// 
        /// date created        20161214
        /// last modified       20161214
        /// author              AGL027
        /// </summary>
        public byte[] GetImage(String wsParmString, String mapOptions)
        {
            /* StringBuilder optionstring = new StringBuilder();
             optionstring.Append(_composeOptionString(options));
             */
            StringBuilder url = new StringBuilder();
            url.Append(_wsns.ToString()).Append("/Image?").Append(wsParmString).Append(_composeMapOptionString(mapOptions)).Append("&api_key=" + _apikey);
            String errorMessage = "";

            byte[] content = _downloadDataFromURL(url.ToString(), ref errorMessage);

            return content;
        }


        /// <summary>
        /// This method with public scope is called by consumer to obtain the map image.  It requires assistance from both  _composeMapOptionString() and _downloadDataFromURL()
        /// 
        /// date created        20171016
        /// last modified       20171018
        /// author              AGL027
        /// </summary>
        public string GetLeafletHTMLSource(String wsParmString, int mapHeight)
        {
            StringBuilder sHtmlSource = new StringBuilder("");

            sHtmlSource.Append(@"
<html>
<head>
    <meta http-equiv='X-UA-Compatible' content='IE = Edge'/>
   <title>AtoBviaC Map with Leaflet</title>
   <link rel='stylesheet' href='leaflet/leaflet.css'/>
   <!--[if lte IE 8]><link rel = 'stylesheet' href = 'leaflet/leaflet.ie.css'/><![endif]-->
       <script src = 'leaflet/leaflet.js' ></script>
         <meta charset = 'utf-8'>
            <meta name = 'viewport' content = 'width =device-width, initial-scale=1.0' >
            <link href = 'docs /images/favicon.ico' rel = 'shortcut icon' type = 'image/x-icon' >
            <link href = 'https://unpkg.com/leaflet@1.2.0/dist/leaflet.css' rel = 'stylesheet' crossorigin = '' integrity = 'sha512-M2wvCLH6DSRazYeZRIm1JnYyh22purTM+FDB5CsyxtQJYeKq83arPe5wgbNmcFXGqiSH2XR8dT/fJISVA1r/zQ=='>
                    <script src = 'https://unpkg.com/leaflet@1.2.0/dist/leaflet.js' crossorigin = '' integrity = 'sha512-lInM/apFSqyy1o6s89K4iQUKg6ppXEgsVxT35HbzUupEVRh2Eu9Wdl4tHj7dZO0s1uvplcYGmt3498TtHq+log=='></script >
                        <script language = 'javascript'>
                        function init() {
            var map = new L.Map('map', {
            	zoomControl: false
		    });                       
                
      	 	L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
	            maxZoom: 18,
	            attribution: '&copy; <a href=""http://www.openstreetmap.org/copyright"">OpenStreetMap</a>'
            }).addTo(map);

            L.tileLayer('https://tiles.openseamap.org/seamark/{z}/{x}/{y}.png', {
	            maxZoom: 18,
	            attribution: 'Map data: &copy; <a href=""http://www.openseamap.org"">OpenSeaMap</a> contributors'
            }).addTo(map);

            map.attributionControl.setPrefix(''); // Don't show the 'Powered by Leaflet' text.");
            
            int countOfLegs = (int)_voyagexml.Descendants(_xmlns + "Leg").Count();

            StringBuilder polylinePoint = new StringBuilder(@"
            //Define an array of Latlng objects (points along the line)
            var polylinePoints = [");
            StringBuilder popupDetail = new StringBuilder();

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
                                EcaZoneToPrevious = r.Element(_xmlns + "EcaZoneToPrevious") != null ? r.Element(_xmlns + "EcaZoneToPrevious").Value : "",
                                LatGeodetic = Convert.ToDecimal(r.Element(_xmlns + "LatGeodetic").Value),
                                Lon = Convert.ToDecimal(r.Element(_xmlns + "Lon").Value)
                            }).ToList()

                        }).ToList();
            foreach (Leg leg in legs)
            {
                foreach (WayPoint wp in leg.WayPointList)
                {
                    if (wp.name!=null && wp.name!="")
                    {
                        popupDetail.Append(@"
                            L.marker([").Append(wp.LatGeodetic.ToString()).Append(",").Append(wp.Lon.ToString()).Append(@"]).addTo(map)
                            .bindPopup('<b>").Append(wp.name).Append("</b><br/>Lat:").Append(String.Format("{0:0.00}", wp.LatGeodetic)).Append(", Lon:").Append(String.Format("{0:0.00}", wp.Lon)).Append(@"')
                            .openPopup();
                        ");

                    }

                    polylinePoint.Append("new L.LatLng(").Append(wp.LatGeodetic.ToString()).Append(",").Append(wp.Lon.ToString()).Append("),");
                }
            }
            sHtmlSource.Append(popupDetail);

            sHtmlSource.Append(polylinePoint);

            sHtmlSource.Append(@"];
         
         var polylineOptions = {
               color: 'green',
               weight: 3,
               opacity: 0.9
             };

         var polyline = new L.Polyline(polylinePoints, polylineOptions);

         map.addLayer(polyline);                        

         // zoom the map to the polyline
         map.fitBounds(polyline.getBounds());

		//add zoom control with your options
		L.control.zoom({
			 position:'topright'
		}).addTo(map);
      }
    </script>
    </head>
    <body onLoad='javascript: init(); '>
        <div id = 'map' style = 'height: ").Append(mapHeight.ToString()).Append(@"px' ></div>
    </body>
</html>
     ");
            return sHtmlSource.ToString();
        }





        /// <summary>
        /// This method is called by the getImage() method to construct the options that can be used when presenting the map image.
        /// 
        /// date created        20161214
        /// last modified       20170912
        /// author              AGL027
        /// </summary>
        private String _composeMapOptionString(String optionMapString)
        {
            List<string> mapOptionList;
            StringBuilder optionQueryString = new StringBuilder();
            mapOptionList = optionMapString.Split('|').ToList();


            if (mapOptionList.Count == 0)
            {
                return "";
            }
            else
            {
                if (mapOptionList[0] == "1")
                {
                    optionQueryString.Append("&zoom=true");
                }
                else
                {
                    optionQueryString.Append("&zoom=false");
                }

                if (mapOptionList[1] == "1")
                {
                    optionQueryString.Append("&showPortLabels=true");
                }
                else
                {
                    optionQueryString.Append("&showPortLabels=false");
                }
                
                /* map height */
                if (mapOptionList[2] != "") optionQueryString.Append("&height=").Append(mapOptionList[2].ToString());
                /* map width */
                if (mapOptionList[3] != "") optionQueryString.Append("&width=").Append(mapOptionList[3].ToString());
                /* port label font color  */
                if (mapOptionList[4] != "") optionQueryString.Append("&portLabelColor=").Append(mapOptionList[4].ToString());
                /* port label font size */
                if (mapOptionList[5] != "") optionQueryString.Append("&portLabelFontSize=").Append(mapOptionList[5].ToString());
                /* coastline background color  */
                if (mapOptionList[6] != "") optionQueryString.Append("&coastlineColor=").Append(mapOptionList[6].ToString());
                /* land background color  */
                if (mapOptionList[7] != "") optionQueryString.Append("&landColor=").Append(mapOptionList[7].ToString());
                /* sea background color  */
                if (mapOptionList[8] != "") optionQueryString.Append("&seaColor=").Append(mapOptionList[8].ToString());

                // We decide to change default colours
                // optionQueryString.Append("&landcolor=35,73,88&coastLineColor=76,188,208");
            }
            return optionQueryString.ToString();
        }

    }




}
