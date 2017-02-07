using System;
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
        private String _voyagestring;
        private XDocument _routingpointxml;
        private XDocument _voyagexml;
        private XDocument _portsxml;
        private StringBuilder _rpShortCodesInLeg = new StringBuilder();
        private StringBuilder _rpNamesInLeg = new StringBuilder();
        private StringBuilder _rpOpenByDefaultInLeg = new StringBuilder();
        private StringBuilder _routingsForDW = new StringBuilder("");

        enum dataFormat
        {
            None,
            XML,
            PBDataWindowSyntax,
            voyageString
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
            _routingpointxml = _getRoutingPoints();
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
        public void setApiKey(String apikey)
        {
            _apikey = apikey;
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
        /// If method is to obtain the map the ports are transformed into port names.
        /// </summary>
        /// 
        /// <remarks>
        /// TODO - expand to handle partial routing points?
        /// 
        /// date created        20161205
        /// last modified       20170119
        /// author              AGL027
        /// </remarks>
        /// <param name="method">possibly could be Image; Voyage</param>
        /// <param name="ports">a delimited list of port codes from Tramos client</param>
        /// <param name="openports">a delimted list of open ports from the Tramos client.  They may or might not include leg index in format SUZ-0|SUZ-1</param>
        /// <param name="closeports">a delimted list of closed ports from the Tramos client.  Optional leg index once again in format DOV-0|DOV-2</param>
        public String transformToUrl(String method, String ports, String openports, String closeports)
        {
            String parturl;
            String portParms = "";
            String openParms = "";
            String closeParms = "";
            String portArg = "";

            List<string> portList;
            List<string> openList;
            List<string> closeList;

            portList = ports.Split('|').ToList();
            openList = openports.Split('|').ToList();
            closeList = closeports.Split('|').ToList();

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

                    if (method.ToLower() == "image")
                    {
                        if (_portsxml == null)
                        {
                            _portsxml = _getPorts();
                        }
                        // obtain nice port name and escape problems chars i.e. <space> 
                        portArg = Uri.EscapeDataString(_getPortName(port));
                    }
                    else
                    {
                        portArg = port;
                    }

                    portParms += "port=" + portArg;

                }
            }
            foreach (var routingpoint in openList)
            {
                if (routingpoint != "")
                {
                    openParms += "&open=" + routingpoint;
                }
            }
            foreach (var routingpoint in closeList)
            {
                if (routingpoint != "")
                {
                    closeParms += "&close=" + routingpoint;
                }
            }

            parturl = portParms + openParms + closeParms;
            return parturl;
        }

        /// <summary>
        /// direct call to return a simple string containing total distance in voyage
        /// 
        /// date created        20161205
        /// last modified       20161205
        /// author              AGL027
        /// </summary>
        public String getDistance(String voyagestring)
        {

            StringBuilder url = new StringBuilder();
            // might be called from within from one of the overridden methods or potentially directly from Tramos
            //_voyagestring = voyagestring;

            url.Append(_wsns.ToString()).Append("/Distance?").Append(voyagestring).Append("&api_key=" + _apikey);

            abcResponse abcResp = new abcResponse();
            abcResp = _downloadStringFromURL(url.ToString());

            if (!string.IsNullOrEmpty(abcResp.httpcode))
            {
                return abcResp.httpcode;
            }
            else
            {
                // _voyagestring = _createNewQryStr(content);
            }
            return abcResp.content;
        }

        /// <summary>
        /// called to transform the calculations stored state.
        /// 
        /// date created        20170124
        /// last modified       20170124
        /// author              AGL027
        /// </summary>
        public String getQueryStringFromAbcEngineState(String ports, String abcEngineState, int scanEca)
        {
            // the plan is to use just the string and the ports to obtain existing items on a leg level, possibly set these as closed.
            // then obtain the open ports by using the full query. replacing or adding new items to list accordingly

            _routingsForDW = new StringBuilder("");

            List<string> portList;
            List<string> openRoutings = new List<string>();
            List<string> closedRoutings = new List<string>();
            List<string> finalRoutings = new List<string>();
            String portParms = "";

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

            urlRoutings.Append(_wsns.ToString()).Append("/Voyage?").Append(portParms).Append("&scaneca=true").Append("&api_key=" + _apikey);
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
            url.Append(_wsns.ToString()).Append("/Voyage?").Append(portParms).Append("&routingString=").Append(abcEngineState).Append("&scaneca=true").Append("&api_key=" + _apikey);

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

            _voyagestring = portParms + "&" + finalRoutingParms + "&scaneca=true";

            System.Windows.Forms.MessageBox.Show(_routingsForDW.ToString());

            return _voyagestring.ToString();
        }

        /// <summary>
        /// returns routings data to be received by datawindow
        /// 
        /// date created        20170131
        /// last modified       20170131
        /// author              AGL027
        /// </summary>
        public String getRoutingsDataForDW()
        {
            return _routingsForDW.ToString();
        }

        /// <summary>
        /// called to transform the calculations stored state.
        /// 
        /// date created        20170124
        /// last modified       20170124
        /// author              AGL027
        /// </summary>
        public String getQueryStringByEngineState(String ports, String abcEngineState, int scanEca, int returnType)
        {


            String portParms = "";

            List<string> openPortList = new List<string>();
            List<string> closedPortList = new List<string>();
            List<string> portList;

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


            /* open routing point parms */
            var rps = from rp in _routingpointxml.Descendants(_xmlns + "RoutingPoint")
                      where rp.Element(_xmlns + "OpenByDefault").Value.Contains("true")
                      select new
                      {
                          RoutingPointCode = rp.Element(_xmlns + "ShortCode").Value,
                      };

            foreach (var rp in rps)
            {
                closedPortList.Add("close=" + rp.RoutingPointCode);
            }

            StringBuilder url = new StringBuilder();

            if (abcEngineState != _voyagestring)
            {

                url.Append(_wsns.ToString()).Append("/Voyage?").Append(portParms).Append("&routingString=").Append(abcEngineState).Append("&scaneca=true").Append("&api_key=" + _apikey);

                abcResponse abcResp = new abcResponse();

                abcResp = _downloadStringFromURL(url.ToString());
                if (!string.IsNullOrEmpty(abcResp.code))
                {
                    return abcResp.code;
                }
                this._voyagexml = XDocument.Parse(abcResp.content);

                _voyagestring = url.ToString();

                abcResp.content = this._voyagexml.ToString();

            }
            // TODO - here we transform the querystring into somthing reasonable

            if (returnType == (int)dataFormat.voyageString)
            {
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
                            openPortList.Add("open=" + wp.routingPointCode + "-" + legIndex.ToString());
                        }
                    }
                    legIndex++;
                }

                // now close all ports on default travel

                string openPortParms = String.Join("&", openPortList.ToArray());
                string closedPortParms = String.Join("&", closedPortList.ToArray());

                _voyagestring = portParms + "&" + closedPortParms + "&" + openPortParms + "&scaneca=true";

                return _voyagestring.ToString();

            }
            return "";
        }

        /// <summary>
        /// called to update the voyage from the client.
        /// 
        /// date created        20161206
        /// last modified       20170131
        /// author              AGL027
        /// </summary>
        public String getVoyage(String voyagestring, int returnType)
        {
            // might be called from within from one of the overridden methods or potentially directly from Tramos

            String content = "";

            if (voyagestring != _voyagestring)
            {
                StringBuilder url = new StringBuilder();
                url.Append(_wsns.ToString()).Append("/Voyage?").Append(voyagestring).Append("&scaneca=true").Append("&api_key=" + _apikey);
                abcResponse abcResp = new abcResponse();
                abcResp = _downloadStringFromURL(url.ToString());
                if (!string.IsNullOrEmpty(abcResp.code))
                {
                    return abcResp.code;
                }
                _voyagestring = voyagestring;
                this._voyagexml = XDocument.Parse(abcResp.content);
            }
            else
            {
                content = this._voyagexml.ToString();
            }

            if (returnType == (int)dataFormat.XML)
            {
                return content;
            }
            else if (returnType == (int)dataFormat.PBDataWindowSyntax)
            {

                /* construct data that will be loaded into the calling datawindow */
                StringBuilder dwContent = new StringBuilder();

                dwContent = _parseXMLContentToDWFormat(content);

                return dwContent.ToString();

            }
            else if (returnType == (int)dataFormat.None)
            {
                return "";
            }
            else if (returnType == (int)dataFormat.voyageString)
            {
                return _voyagestring;
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
                            wayPointData.Append("").Append("\t").Append(wp.name).Append("\t").Append(sumOfDistancesSinceLastKnown.ToString()).Append("\t").Append(sumOfEcaZoneDistancesSinceLastKnown.ToString()).Append("\t").Append(is_eca).Append("\r\n");

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
        /// date created        20161209
        /// last modified       20170201
        /// author              AGL027
        /// </summary>
        public String getRoutingPointsForSelectedLeg(int journeyId, String voyagestring)
        {
            _rpShortCodesInLeg.Clear();
            _rpNamesInLeg.Clear();
            _rpOpenByDefaultInLeg.Clear();

            _routingsForDW = new StringBuilder("");

            if (voyagestring != _voyagestring)
            {
                String returnCode = this.getVoyage(voyagestring, 0);
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
        /// date created        ??????
        /// last modified       ??????
        /// author              AGL027
        /// </summary>
        public String getRPNamesInsideLeg()
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
        public String getRPShortCodesInsideLeg()
        {
            return _rpShortCodesInLeg.ToString();
        }

        /// <summary>
        /// Obtain single delimited collection of routing point short codes that are open inside a string
        /// 
        /// date created        ??????
        /// last modified       ??????
        /// author              AGL027
        /// </summary>
        public String getRPOpenByDefaultInsideLeg()
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
            var rpname = (from rp in _routingpointxml.Descendants(_xmlns + "RoutingPoint")
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
            var portname = (from p in _portsxml.Descendants(_xmlns + "Port")
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
            var rps = (from rp in _routingpointxml.Descendants(_xmlns + "RoutingPoint")
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
        public string getRPfromQueryStringInLeg(int journeyId, String rpType)
        {

            var parmData = HttpUtility.ParseQueryString(this._voyagestring).Get(rpType.ToLower());
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
        /// last modified       20170206
        /// author              AGL027
        /// </summary>
        public string getRoutingDWfromQryStr()
        {
            System.Windows.Forms.MessageBox.Show("Gets here!");

            _routingsForDW = new StringBuilder("");

            var parmData = HttpUtility.ParseQueryString(this._voyagestring);

            foreach (string open in parmData["open"].Split(','))
            {
                _routingsForDW.Append(open).Append("\t").Append(_getRPName(open.Substring(0, 3))).Append("\t").Append("1").Append("\t").Append(_getOpenByDefault(open.Substring(0, 3))).Append("\r\n");
            }
            
            foreach (string close in parmData["close"].Split(','))
            {
                _routingsForDW.Append(close).Append("\t").Append(_getRPName(close.Substring(0, 3))).Append("\t").Append("0").Append("\t").Append(_getOpenByDefault(close.Substring(0, 3))).Append("\r\n");
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
        public String getRPbyType(int rpType, String queryParm)
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

            var rps = from rp in _routingpointxml.Descendants(_xmlns + "RoutingPoint")
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
        public byte[] getImage(String voyagestring, String mapOptions)
        {
            /* StringBuilder optionstring = new StringBuilder();
             optionstring.Append(_composeOptionString(options));
             */
            StringBuilder url = new StringBuilder();

            url.Append(_wsns.ToString()).Append("/Image?").Append(voyagestring).Append(_composeMapOptionString(mapOptions)).Append("&api_key=" + _apikey);

            String errorMessage = "";

            byte[] content = _downloadDataFromURL(url.ToString(), ref errorMessage);

            return content;
        }

        /// <summary>
        /// This method is called by the getImage() method to construct the options that can be used when presenting the map image.
        /// 
        /// date created        20161214
        /// last modified       20161214
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
                /* AtoBviaC default for SECA scanZones is False */
                if (mapOptionList[0] == "1")
                {
                    optionQueryString.Append("&showSecaZones=true");
                }
                else
                {
                    optionQueryString.Append("&showSecaZones=false");
                }

                if (mapOptionList[1] == "1")
                {
                    optionQueryString.Append("&zoom=true");
                }
                else
                {
                    optionQueryString.Append("&zoom=false");
                }

                /* AtoBviaC default for AntiPiracy is True */
                if (mapOptionList[2] != "") optionQueryString.Append("&height=").Append(mapOptionList[2].ToString());
                /* AtoBviaC default for Environmental/Navigational/Regulatory is True */
                if (mapOptionList[3] != "") optionQueryString.Append("&width=").Append(mapOptionList[3].ToString());

                // We need to decide if we change default colours
                // optionQueryString.Append("&landcolor=35,73,88&coastLineColor=76,188,208");
            }
            return optionQueryString.ToString();
        }

    }




}
