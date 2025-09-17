using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
using System.CodeDom.Compiler;
using System.Diagnostics;
using System.Xml;
using System.Xml.Linq;
using Gpx;
namespace Rando
{
    public partial class Rando : Form
    {
        public Rando()
        {
            InitializeComponent();
        }

        List<Trackpoint> trackpoints = new List<Trackpoint>();

        private void Rando_Form_Paint(object sender, PaintEventArgs e)
        {
            trackpoints = DoThat();
            if (trackpoints.Count < 2)
                return;

            // Récupérer min/max
            double minLat = trackpoints.Min(p => p.Latitude);
            double maxLat = trackpoints.Max(p => p.Latitude);
            double minLon = trackpoints.Min(p => p.Longitude);
            double maxLon = trackpoints.Max(p => p.Longitude);

            // Marge pour éviter que ça touche les bords
            int margin = 20;
            int width = this.ClientSize.Width - 2 * margin;
            int height = this.ClientSize.Height - 2 * margin;

            // Conversion GPS -> pixels
            Point[] points = trackpoints.Select(p =>
            {
                int x = margin + (int)((p.Longitude - minLon) / (maxLon - minLon) * width);
                int y = margin + (int)((maxLat - p.Latitude) / (maxLat - minLat) * height); // inversion Y
                return new Point(x, y);
            }).ToArray();

            Pen myPen = new Pen(Color.Red);
            myPen.Width = 2;

            //Point[] points = new Point[4] { new Point(30, 50), new Point(50, 10), new Point(80, 50), new Point(111, 400) };
            this.CreateGraphics().DrawLines(myPen, points);
        }

        private List<Trackpoint> DoThat()
        {
            // Fonction qui lit le document gpx, qui trouve les trkpt, les valeurs de lat, lon et ele et les associe aux attributs de la classe Trackpoint
            //XDocument doc = XDocument.Load("gemmikandersteg.gpx");

            var input = new StreamReader("gemmikandersteg.gpx");
            using (GpxReader reader = new GpxReader(input.BaseStream))
            {

                while (reader.Read())
                {
                    switch (reader.ObjectType)
                    {
                        case GpxObjectType.Metadata:
                            //writer.WriteMetadata(reader.Metadata);
                            break;
                        case GpxObjectType.WayPoint:
                            //writer.WriteWayPoint(reader.WayPoint);
                            break;
                        case GpxObjectType.Route:
                            //writer.WriteRoute(reader.Route);
                            break;
                        case GpxObjectType.Track:
                            //writer.WriteTrack(reader.Track);
                            var trackInfo = reader.Track;

                            var pointsFromLib = trackInfo.ToGpxPoints();

                            trackpoints.AddRange(pointsFromLib.Select(gpx=>new Trackpoint() { Latitude=gpx.Latitude, Elevation=gpx.Elevation, Longitude=gpx.Longitude }));

                            //TODO : lecuperer les points de track et les convertir vers trackpoint
                            break;
                    }
                }

            }
            return trackpoints;

        }


    }
    class Trackpoint
    {
        private double _latitude;
        private double _longitude;
        private double? _elevation;

        public double Latitude { get => _latitude; set => _latitude = value; }
        public double Longitude { get => _longitude; set => _longitude = value; }
        public double? Elevation { get => _elevation; set => _elevation = value; }
    }
}

