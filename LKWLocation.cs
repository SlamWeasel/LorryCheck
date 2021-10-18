using System;

namespace Lastwagen_Abfrage
{
    class LKWLocation
    {
        public LKWLocation()
        {

        }

        public string LKW { get; set; }
        public DateTime LastUpdatedUtc { get; set; }
        public string Position { get; set; }
        public float Latitude { get; set; }
        public float Longitude { get; set; }

        public override string ToString()
            => $"{LKW}, {LastUpdatedUtc}, {Position}, {Latitude}, {Longitude}";
    }
}
