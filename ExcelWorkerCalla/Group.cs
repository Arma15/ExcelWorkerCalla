using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWorkerCalla
{
    #region Group Class
    public class Group
    {
        public List<string> geoNumbers
        {
            get
            {
                if (balls.Count < 1)
                {
                    return new List<string>();
                }

                foreach (GeometryData geo in balls[0].geometries)
                {
                    _geoNumbers.Add(geo.geoNumber);
                }
                return _geoNumbers;
            }
        }

        public string GroupNumber
        {
            get
            {
                return _groupNum;
            }
        }

        private string _groupNum;
        private List<string> _geoNumbers;
        public List<Ball> balls;

        public Group()
        {
            balls = new List<Ball>();
            _geoNumbers = new List<string>();
        }

        public Group(string groupNum)
        {
            _groupNum = groupNum;
            balls = new List<Ball>();
            _geoNumbers = new List<string>();
        }

        public Group(List<Ball> newBalls)
        {
            balls = new List<Ball>(newBalls);
            _geoNumbers = new List<string>();
        }

        public void Add(Ball ball)
        {
            balls.Add(ball);
        }

        public double[] AveGeometryFields(string geoNum)
        {
            double[] totals = new double[11];
            for (int i = 0; i < totals.Length; ++i)
            {
                totals[i] = 0.0;
            }

            foreach (Ball b in balls)
            {
                GeometryData gd = b.FindGeometry(geoNum);
                totals[0] += gd.height;
                totals[1] += gd.width;
                totals[2] += gd.totalArea;
                totals[3] += gd.areaTop;
                totals[4] += gd.flatness;
                totals[5] += gd.maxCurvature;
                totals[6] += gd.maxSlopeAve;
                totals[7] += gd.maxSlopeXAve;
                totals[8] += gd.maxSlopeRAve;
                totals[9] += gd.slopeWidth;
                totals[10] += gd.recirculationAreaAve;
            }

            // Average totals
            for (int i = 0; i < totals.Length; ++i)
            {
                totals[i] = totals[i] / balls.Count; ;
            }

            return totals;
        }
    }
    #endregion
}
