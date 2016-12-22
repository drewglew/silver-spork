using System;
using System.Net;
using System.Text;
using System.Runtime.InteropServices;
using System.Xml.Serialization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;

namespace mt.distances
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
        public String updateVoyage(String voyagestring, int returnxml)
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

            if (returnxml == 1)
            {
                return content;
            }
            else
            {
                return "";
            }


        }


        /* 
        * date created 20161206
        * last modified 20161215
        * unproven, but called to update the voyage from the client 
        */
        public String getVoyage(String voyagestring, int sendxml)
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

            if (sendxml == 1)
            {
                return content;
            }
            else
            {
                return "";
            }


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
            
            if (voyagestring!=_voyagestring)
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
                       where  voyage.Element(_xmlns + "ToPort").Element(_xmlns + "Code").Value.Contains(toPort) && voyage.Element(_xmlns + "FromPort").Element(_xmlns + "Code").Value.Contains(fromPort)
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
            var rps = (from rp in _routingpointxml.Descendants(_xmlns + "RoutingPoint")
                      where rp.Element(_xmlns + "ShortCode").Value.Contains(rpShortCode)
                      select rp.Element(_xmlns + "Name").Value).FirstOrDefault();
            return rps;
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
        public String getRPbyType(int rpType,String queryParm)
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
}
