using System;
using System.Net;
using System.Text;
using System.Runtime.InteropServices;
using System.Xml.Serialization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;
using tankers.distances.models.datawindow;
using System.Collections;

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
        private StringBuilder _rpShortCodesInLeg = new StringBuilder();
        private StringBuilder _rpNamesInLeg = new StringBuilder();
        private StringBuilder _rpOpenByDefaultInLeg = new StringBuilder();

        enum dataFormat
        {
            None,
            XML,
            PBDataWindowSyntax
        };

        /* date created 20161205 */
        /* last modified 20161205 */
        /* constructor */
        public AtoBviaC()
        {
            _xmlns = "http://api.atobviaconline.com/v1";
            _wsns = "https://api.atobviaconline.com/v1";
            _routingpointxml = _getRoutingPoints();
        }

        /* date created 20161215 */
        /* last modified 20161215 */
        /* No need to hardcode the api_key value in here. */
        public void setApiKey(String apikey)
        {
            _apikey = apikey;
        }

        /* date created 20161205 
         * last modified 20161205 
         * low level communication to web service; this is to handle image processing */
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
                if (ex.Status == WebExceptionStatus.ProtocolError && ex.Response != null)
                {
                    errorMessage = "HttpWebResponse: " + (int)resp.StatusCode + ":" + resp.StatusCode.ToString();
                    System.Windows.Forms.MessageBox.Show(errorMessage.ToString());
                }
            }
            return content;
        }

        /* date created 20161205 
         * last modified 20161205 
         * low level communication to web service; this is to handle strings */
        private String _downloadStringFromURL(String url, ref String errorMessage)
        {
            String content = "";

            try
            {
                using (WebClient wc = new WebClient())
                {
                    wc.Headers.Add("accept", "application/xml");
                    content = wc.DownloadString(url);
                }

            }
            catch (WebException ex)
            {
                var resp = (HttpWebResponse)ex.Response;
                if (ex.Status == WebExceptionStatus.ProtocolError && ex.Response != null)
                {
                    errorMessage = "HttpWebResponse: " + (int)resp.StatusCode + ":" + resp.StatusCode.ToString();
                }
            }
            return content;
        }

        /* date created 20161205
         * last modified 20161206
         * works 
         * TODO - expand to handle partial routing points?
         */
        public String transformToUrl(String method, String ports, String openports, String closeports)
        {

            /* taken away the routing string temporarily.
             * this method transforms the delimited items received from PowerBuilder 
             * and transforms into some sort of querystring segment.
             */

            String parturl;
            String portParms = "";
            String openParms = "";
            String closeParms = "";

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
                    portParms += "port=" + port;
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



        /* date created 20161205 
         * last modified 20161205
         * proven to work - direct call to return a simple string containing total distance in voyage */
        public String getDistance(String voyagestring)
        {

            StringBuilder url = new StringBuilder();
            // might be called from within from one of the overridden methods or potentially directly from Tramos
            //_voyagestring = voyagestring;

            url.Append(_wsns.ToString()).Append("/Distance?").Append(voyagestring).Append("&api_key=" + _apikey);
            String errorMessage = "";
            String content = _downloadStringFromURL(url.ToString(), ref errorMessage);

            if (errorMessage != "")
            {
                return errorMessage + "<url=" + url.ToString() + ">";
            }
            else
            {
                // _voyagestring = _createNewQryStr(content);
            }
            return content;
        }


        /* 
        * date created 20161219
        * last modified 20161219
        * called by getRoutingPointsForSelectedLeg() validates existing querystring for voyage against the one past in.
        */
        public String updateVoyage(String voyagestring, int returnType)
        {
            // might be called from within from one of the overridden methods or potentially directly from Tramos

            String content = "";

            if (voyagestring != _voyagestring)
            {
                _voyagestring = voyagestring;

                StringBuilder url = new StringBuilder();

                url.Append(_wsns.ToString()).Append("/Voyage?").Append(voyagestring).Append("&scaneca=true").Append("&api_key=" + _apikey);

                String errorMessage = "";
                content = _downloadStringFromURL(url.ToString(), ref errorMessage);
                this._voyagexml = XDocument.Parse(content);
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
            return "error, unhandled";

        }


        /* 
        * date created 20161206
        * last modified 20161215
        * unproven, but called to update the voyage from the client 
        */
        public String getVoyage(String voyagestring, int returnType)
        {
            // might be called from within from one of the overridden methods or potentially directly from Tramos

            String content = "";

            if (voyagestring != _voyagestring)
            {
                _voyagestring = voyagestring;

                StringBuilder url = new StringBuilder();

                url.Append(_wsns.ToString()).Append("/Voyage?").Append(voyagestring).Append("&scaneca=true").Append("&api_key=" + _apikey);

                String errorMessage = "";
                content = _downloadStringFromURL(url.ToString(), ref errorMessage);
                this._voyagexml = XDocument.Parse(content);
            } else
            {
                content = this._voyagexml.ToString();
            }

            if (returnType == (int)dataFormat.XML)
            {
                return content;
            } else if (returnType == (int)dataFormat.PBDataWindowSyntax)
            {

                /* construct data that will be loaded into the calling datawindow */
                StringBuilder dwContent = new StringBuilder();

                dwContent = _parseXMLContentToDWFormat(content);

                return dwContent.ToString();

            } else if (returnType == (int)dataFormat.None)
            {
                return "";
            }
            return "error, unhandled";
        }


        /* 
        * date created 20170116
        * last modified 20170118
        * Obtain delimited routing data from the voyage XML string
        * DK0021~tCopenhagen~t0.000~t0.000~t1\r~nSOU~tRP Name~t23.223~t23.223~t1\r\n~tExit Baltic zone, Enter North Sea zone~t121.493~t144.716~t1\r~nSKA~tRP Name~t134.833~t13.340~t1\r~nGB0294~tLondon~t591.158~t604.498\t~t1\r\n
        */
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

            foreach (Leg leg in legs)
            {
                decimal distancePorttoPort = 0;
                decimal sumOfDistancesSinceLastKnown = 0;
                decimal LastPortDistanceFromStart = 0;
                decimal sumOfEcaZoneDistancesSinceLastKnown = 0;

               // string EcaZoneLastPort = "";

                int wp_counter = 0;
               // bool is_eca = false;
                int is_eca = 0;
                
                /* now the fun begins */
                foreach (WayPoint wp in leg.WayPointList)
                {
                    if (wp_counter==0)
                    {
                        if (wp.EcaZoneToPrevious!="")
                        {
                            // we start inside an eca zone
                            leg.start_in_eca = true;
                            is_eca = 1;
                        } else
                        {
                            leg.start_in_eca = false;
                            is_eca = 0;
                        }
                    }

                    distancePorttoPort = wp.DistanceFromStart - LastPortDistanceFromStart;
                    sumOfDistancesSinceLastKnown += distancePorttoPort;

                    if (is_eca == 1)
                    {
                        sumOfEcaZoneDistancesSinceLastKnown += distancePorttoPort;
                    }

                    if (leg.fromPort.code == wp.name)
                    {
                        // we are in the first port
                        fromPortData.Append(leg.fromPort.code).Append("\t").Append(leg.fromPort.name).Append("\t").Append("0.000").Append("\t").Append("0.000").Append("\t").Append(is_eca).Append("\r\n");
                    } else if (leg.toPort.code == wp.name)
                    {
                        // we are in the last port
                        toPortData.Append(leg.toPort.code).Append("\t").Append(leg.toPort.name).Append("\t").Append(sumOfDistancesSinceLastKnown.ToString()).Append("\t").Append(sumOfEcaZoneDistancesSinceLastKnown.ToString()).Append("\t").Append(is_eca).Append("\r\n");
                    }
                    else
                    {
                        if (wp.name.Length>4 &&  wp.name.Substring(0, 5) == "Exit ")
                        {
                            wayPointData.Append("").Append("\t").Append(wp.name).Append("\t").Append(sumOfDistancesSinceLastKnown.ToString()).Append("\t").Append(sumOfEcaZoneDistancesSinceLastKnown.ToString()).Append("\t").Append(is_eca).Append("\r\n");
                        }
                        else if (wp.routingPointCode != "")
                        {
                            wayPointData.Append(wp.routingPointCode).Append("\t").Append(_getRPName(wp.routingPointCode)).Append("\t").Append(sumOfDistancesSinceLastKnown.ToString()).Append("\t").Append(sumOfEcaZoneDistancesSinceLastKnown.ToString()).Append("\t").Append(is_eca).Append("\r\n");
                            sumOfDistancesSinceLastKnown = 0;
                        }

                        if (wp.name.Length > 4 && wp.name.Substring(0, 5) == "Exit ")
                        {
                            sumOfEcaZoneDistancesSinceLastKnown = 0;
                        }

                    }
                    LastPortDistanceFromStart = wp.DistanceFromStart;
                    wp_counter++;
                }

                
            }
            dwContent.Append(fromPortData).Append(wayPointData).Append(toPortData);

            return dwContent;
        }


        /* 
        * date created 20161209
        * last modified 20161222
        * brand new - to do - perhaps work a better means to share the name vars of routing points
        */
        public String getRoutingPointsForSelectedLeg(int journeyId, String voyagestring)
        {
            // might be called from within from one of the overridden methods or potentially directly from Tramos

            _rpShortCodesInLeg.Clear();
            _rpNamesInLeg.Clear();
            _rpOpenByDefaultInLeg.Clear();

            if (voyagestring != _voyagestring)
            {
                this.updateVoyage(voyagestring, 0);
            }


            var leg = from voyage in _voyagexml.Descendants(_xmlns + "Leg").Skip(journeyId - 1).Take(1)
                      select voyage.Element(_xmlns + "Waypoints");

            StringBuilder routingPointsWithinLeg = new StringBuilder();

            foreach (var item in leg.Elements(_xmlns + "Waypoint"))
            {
                String rpShortCode = item.Element(_xmlns + "RoutingPoint") != null ? item.Element(_xmlns + "RoutingPoint").Value : "";

                if (rpShortCode != "")
                {
                    _rpShortCodesInLeg.Append(rpShortCode + "|");
                    _rpNamesInLeg.Append(_getRPName(rpShortCode) + "|");
                    _rpOpenByDefaultInLeg.Append(_getOpenByDefault(rpShortCode) + "|");
                }

            }
            return _rpShortCodesInLeg.ToString();
        }


        /* 
        * date created 20161209
        * last modified 20161212
        * brand new - now it works!
        */
        public String getRPsInsideVoyageLeg(int journeyId, String fromPort, String toPort)
        {
            // might be called from within from one of the overridden methods or potentially directly from Tramos

            _rpShortCodesInLeg.Clear();
            _rpNamesInLeg.Clear();
            _rpOpenByDefaultInLeg.Clear();

            var leg = from voyage in _voyagexml.Descendants(_xmlns + "Leg")
                      where voyage.Element(_xmlns + "ToPort").Element(_xmlns + "Code").Value.Contains(toPort) && voyage.Element(_xmlns + "FromPort").Element(_xmlns + "Code").Value.Contains(fromPort)
                      select voyage.Element(_xmlns + "Waypoints");

            StringBuilder routingPointsWithinLeg = new StringBuilder();

            foreach (var item in leg.Elements(_xmlns + "Waypoint"))
            {
                String rpShortCode = item.Element(_xmlns + "RoutingPoint") != null ? item.Element(_xmlns + "RoutingPoint").Value : "";

                if (rpShortCode != "") {
                    _rpShortCodesInLeg.Append(rpShortCode + "|");
                    _rpNamesInLeg.Append(_getRPName(rpShortCode) + "|");
                    _rpOpenByDefaultInLeg.Append(_getOpenByDefault(rpShortCode) + "|");
                }

            }

            return _rpShortCodesInLeg.ToString();
        }


        public String getRPNamesInsideLeg()
        {
            return _rpNamesInLeg.ToString();
        }

        public String getRPShortCodesInsideLeg()
        {
            return _rpShortCodesInLeg.ToString();
        }

        public String getRPOpenByDefaultInsideLeg()
        {
            return _rpOpenByDefaultInLeg.ToString();
        }


        /* 
        * date created 20161206
        * last modified 20161206
        * unproven to work - currently works without the api-key as this is called before api-key is set.
        */
        private XDocument _getRoutingPoints()
        {
            String url = "https://api.atobviaconline.com/v1/RoutingPoints?api_key=" + _apikey;
            WebClient syncClient = new WebClient();
            syncClient.Headers.Add("accept", "application/xml");
            String content = syncClient.DownloadString(url);
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

        /* 
        * created date  20161213
        * last modified 20161213
        */
        private String _getRPName(String rpShortCode)
        {
            var rpname = (from rp in _routingpointxml.Descendants(_xmlns + "RoutingPoint")
                       where rp.Element(_xmlns + "ShortCode").Value.Contains(rpShortCode)
                       select rp.Element(_xmlns + "Name").Value).FirstOrDefault();
            return rpname;
        }

        private String _getOpenByDefault(String rpShortCode)
        {
            var rps = (from rp in _routingpointxml.Descendants(_xmlns + "RoutingPoint")
                       where rp.Element(_xmlns + "ShortCode").Value.Contains(rpShortCode)
                       select rp.Element(_xmlns + "OpenByDefault").Value).FirstOrDefault();
            return rps;
        }

        /* 
        * created date  20161129
        * last modified 20161208
        */
        public String getRPbyType(int rpType, String queryParm)
        {
            /*  getRPbyType - AtoBviaC have a set of default Open and Closed ports.  This method
                provides a querystring segment that is used on first presentation of routing 
            */

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

        /* 
        * created date  20161214
        * last modified 20161214
        * comment       new
        */
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

        /* 
        * created date  20161214
        * last modified 20161214
        * comment       new
        */
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

                optionQueryString.Append("&landcolor=35,73,88&coastLineColor=76,188,208");


            }
            return optionQueryString.ToString();
        }

        
    }

    public class Port
    {
        public string name { get; set; }
        public string code { get; set; }
    }

    public class Leg
    {
        public Port fromPort { get; set; }
        public Port toPort { get; set; }
        public decimal distance { get; set; }
        public decimal eca_distance { get; set; }
        public bool start_in_eca { get; set; }
        public List<WayPoint> WayPointList { get; set;}
    }

    public class WayPoint
    {
        public string name { get; set; }
        public string routingPointCode { get; set; }
        public decimal DistanceFromStart { get; set; }
        public string EcaZoneToPrevious { get; set; }
    }

    
}
