using System;
using System.Collections.Generic;

namespace tankers.distances.models.datawindow
{
    public class Port
    {
        public string name { get; set; }
        public string code { get; set; }
        public double LatGeodetic { get; set; }
        public double Lon { get; set; }
    }

    public class Leg
    {
        public Port fromPort { get; set; }
        public Port toPort { get; set; }
        public decimal distance { get; set; }
        public decimal eca_distance { get; set; }
        public bool start_in_eca { get; set; }
        public List<WayPoint> WayPointList { get; set; }
     
    }

    public class WayPoint
    {
        public string name { get; set; }
        public string routingPointCode { get; set; }
        public decimal DistanceFromStart { get; set; }
        public string EcaZoneToPrevious { get; set; }
        public double LatGeodetic { get; set; }
        public double Lon { get; set; }
    }

    public class ActiveRoutingPoint
    {
        public string Type { get; set; }
        public string ShortCode { get; set; }
        public string Name { get; set; }
        public double LatGeodetic { get; set; }
        public double Lon { get; set; }
        public int LegIndex { get; set; }
        public bool IsOpen { get; set; }
        public bool IsAdvanced { get; set; }

    }
}

