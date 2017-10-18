using System;
using System.Collections.Generic;

namespace tankers.distances.models.datawindow
{
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
        public List<WayPoint> WayPointList { get; set; }
    }

    public class WayPoint
    {
        public string name { get; set; }
        public string routingPointCode { get; set; }
        public decimal DistanceFromStart { get; set; }
        public string EcaZoneToPrevious { get; set; }
        public decimal LatGeodetic { get; set; }
        public decimal Lon { get; set; }
    }




}

