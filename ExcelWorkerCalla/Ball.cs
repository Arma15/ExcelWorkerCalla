using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWorkerCalla
{
    #region Geometry structure
    public struct GeometryData
    {
        // 21-22
        public string geoNumber;
        public string geoNumberIco;
        public string hemisphere;
        public int icosahedron;
        public string geoType;
        public double height;
        public double width;
        public double totalArea;
        public double areaTop;
        public double flatness;
        public double maxCurvature;
        public double maxSlopeAve;
        public double maxSlopeXAve;
        public double maxSlopeRAve;
        public double slopeWidth;
        public double recirculationAreaAve;
    }
    #endregion

    #region Ball Class
    public class Ball
    {
        public string BallNumber
        {
            get
            {
                return _ballNum;
            }
        }
        private string _ballNum;
        public string GroupNumber
        {
            get
            {
                return _groupNum;
            }
        }
        private string _groupNum;
        public List<GeometryData> geometries;

        public Ball(string groupNum, string ballNum)
        {
            _ballNum = ballNum;
            _groupNum = groupNum;
            geometries = new List<GeometryData>();
        }

        /// <summary>
        /// Adds a new geometry to the list for this ball object
        /// </summary>
        /// <param name="gd">Geometry data to be added to this specific ball</param>
        public void AddGeometry(GeometryData gd)
        {
            geometries.Add(gd);
        }

        /// <summary>
        /// Finds the specified geometry number and returns the corresponding object
        /// </summary>
        /// <param name="geoNum">The geometry number to be found</param>
        /// <returns></returns>
        public GeometryData FindGeometry(string geoNum)
        {
            foreach (GeometryData gd in geometries)
            {
                if (gd.geoNumber == geoNum.Trim())
                {
                    return gd;
                }
            }
            return new GeometryData();
        }

        public List<GeometryData> FindAllTop()
        {
            List<GeometryData> top = new List<GeometryData>();

            foreach (GeometryData gd in geometries)
            {
                if (gd.hemisphere.ToLower() == "top")
                {
                    top.Add(gd);
                }
            }
            return top;
        }

        public List<GeometryData> FindAllBottom()
        {
            List<GeometryData> bottom = new List<GeometryData>();

            foreach (GeometryData gd in geometries)
            {
                if (gd.hemisphere.ToLower() == "bottom")
                {
                    bottom.Add(gd);
                }
            }

            return bottom;
        }

        public List<GeometryData> FindByIcosahedron(int num)
        {
            List<GeometryData> icos = new List<GeometryData>();

            foreach (GeometryData gd in geometries)
            {
                if (gd.icosahedron == num)
                {
                    icos.Add(gd);
                }
            }

            return icos;
        }

        public double AveHeight()
        {
            double total = 0.0;
            foreach (GeometryData gd in geometries)
            {
                total += gd.height;
            }

            return total / geometries.Count;
        }

        public double AveWidth()
        {
            double total = 0.0;
            foreach (GeometryData gd in geometries)
            {
                total += gd.width;
            }

            return total / geometries.Count;
        }

        public double AveAreaTotal()
        {
            double total = 0.0;
            foreach (GeometryData gd in geometries)
            {
                total += gd.totalArea;
            }

            return total / geometries.Count;
        }

        public double AveAreaTop()
        {
            double total = 0.0;
            foreach (GeometryData gd in geometries)
            {
                total += gd.areaTop;
            }

            return total / geometries.Count;
        }

        public double AveFlatness()
        {
            double total = 0.0;
            foreach (GeometryData gd in geometries)
            {
                total += gd.flatness;
            }

            return total / geometries.Count;
        }

        public double AveMaxCurvature()
        {
            double total = 0.0;
            foreach (GeometryData gd in geometries)
            {
                total += gd.maxCurvature;
            }

            return total / geometries.Count;
        }

        public double AveMaxSlope()
        {
            double total = 0.0;
            foreach (GeometryData gd in geometries)
            {
                total += gd.maxSlopeAve;
            }

            return total / geometries.Count;
        }

        public double AveMaxSlopeX()
        {
            double total = 0.0;
            foreach (GeometryData gd in geometries)
            {
                total += gd.maxSlopeXAve;
            }

            return total / geometries.Count;
        }

        public double AveMaxSlopeR()
        {
            double total = 0.0;
            foreach (GeometryData gd in geometries)
            {
                total += gd.maxSlopeRAve;
            }

            return total / geometries.Count;
        }

        public double AveSlopeWidth()
        {
            double total = 0.0;
            foreach (GeometryData gd in geometries)
            {
                total += gd.slopeWidth;
            }

            return total / geometries.Count;
        }

        public double AveRecirculationArea()
        {
            double total = 0.0;
            foreach (GeometryData gd in geometries)
            {
                total += gd.recirculationAreaAve;
            }

            return total / geometries.Count;
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();

            sb.AppendLine($"Ball Number: {_ballNum}");
            sb.AppendLine($"Average Height = {AveHeight()}");
            sb.AppendLine($"Average Width = {AveWidth()}");
            sb.AppendLine($"Average Area Total = {AveAreaTotal()}");
            sb.AppendLine($"Average Area Top = {AveAreaTop()}");
            sb.AppendLine($"Average Flatness = {AveFlatness()}");
            sb.AppendLine($"Average Max Curvature = {AveMaxCurvature()}");
            sb.AppendLine($"Average Max Slope Ave = {AveMaxSlope()}");
            sb.AppendLine($"Average Max Slope X Ave = {AveMaxSlopeX()}");
            sb.AppendLine($"Average Max Slope R Ave = {AveMaxSlopeR()}");
            sb.AppendLine($"Average Slope Width = {AveSlopeWidth()}");
            sb.AppendLine($"Average Recirculation Area Avg = {AveRecirculationArea()}");

            return sb.ToString();
        }
    }
    #endregion
}
