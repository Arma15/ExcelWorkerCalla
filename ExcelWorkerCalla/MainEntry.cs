using log4net;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace ExcelWorkerCalla
{
    class MainEntry
    {
        private static readonly ILog _log = LogManager.GetLogger("ExcelWorkerCalla.log");
        public static List<string> geoNumbers = new List<string>();
        public static Group currGroup = new Group();

        static void Main(string[] args)
        {
            // Given a path to a list of ball excel files (param 1)
            string ballPath = @"C:\Users\kflor\OneDrive\Desktop\Archive\Reports\20200319_Group1_Stage1";
            // Given path to the report template to be copied to desired location (param 2)
            string excelPath = @"C:\Users\kflor\OneDrive\Desktop\Archive\Reports\WorkOrder Report Template.xlsx";
            // Given work order number (param 3)
            string workOrderNumber = "0012B";

            for (int i = 0; i < args.Length; ++i)
            {
                _log.Info($"Param #{i + 1}: {args[i]}");
            }

            if (args.Length < 3)
            {
                _log.Error($"Only {args.Length} arguments passed in, expecting 3 parameters");
                // return;
            }

            if (args.Length != 0)
            {
                ballPath = args[0];
                excelPath = args[1];
                workOrderNumber = args[2];
            }

            #region Removed Code
            /*
            double[] ave30Tubes = new double[11];
            double[] aveTopTubes = new double[11];
            double[] aveBottomTubes = new double[11];
            */
            #endregion
            // Validate both directories
            if (!Directory.Exists(ballPath))
            {
                _log.Error($"Invalid directory {ballPath}");
                return;
            }

            if (!File.Exists(excelPath))
            {
                _log.Error($"Invalid path {excelPath}");
                return;
            }

            // Get reports folder path
            string reportsFolderDir = Path.GetDirectoryName(ballPath);

            // Check if directory exists
            if (!Directory.Exists(reportsFolderDir))
            {
                _log.Error($"Directory: {reportsFolderDir} does not exist.");
            }
            
            // Get Work Order reports path
            string workOrderFolder = reportsFolderDir + "\\02_WorkOrder Reports";

            // Check if the work order folder exists
            if (!Directory.Exists(workOrderFolder))
            {
                _log.Error($"Directory: {workOrderFolder} does not exist, creating it now...");
                Directory.CreateDirectory(workOrderFolder);
            }

            // Copy template to workOrderFolder and rename it
            // Get aspects of file name, 0 = date, 1 = time, 2 = group#, 3 = stage#
            string[] info = Path.GetFileName(ballPath).Split('_');
            string newReportFilePath = workOrderFolder + "\\" + info[0] + "_" + info[1] + "_" + workOrderNumber + "_Report.xlsm";
            try
            {
                File.Copy(excelPath, newReportFilePath);
            }
            catch (Exception ex)
            {
                _log.Error($"Error when copying template file: {ex.Message.ToString()}");
                return;
            }

            // Now can write to new report file: newReportFilePath

            try
            {
                DirectoryInfo direct = new DirectoryInfo(ballPath);
                FileInfo[] files = direct.GetFiles("*.xlsx");
                ReadData(files);
            }
            catch (Exception ex)
            {
                _log.Error($"Exception thrown : {ex.Message.ToString()}");
            }
            
            // **************** Test code **********************************
            using (StreamWriter sw = new StreamWriter(@"C:\Users\kflor\OneDrive\Desktop\Averages.txt"))
            {
                // Enter required data to textfile
                foreach (Ball ball in currGroup.balls)
                {
                    sw.Write(ball.ToString());
                    sw.WriteLine();
                }
            }
            // **************** End test code ******************************

            // Grab averages from every geometry
            // AveGeometryFields(string geoNumber)

            // Open excel template file
            // Create a new Excel package from the file
            FileInfo fi = new FileInfo(newReportFilePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excelDoc = new ExcelPackage(fi))
            {
                // open sheet with group number as the name
                //Get first WorkSheet. Note that EPPlus indexes start at 1!
                ExcelWorksheet firstWorksheet = excelDoc.Workbook.Worksheets["Group2"];

                // Find line to insert group number on
                int counter = 1;
                while (firstWorksheet.Cells[counter, 1].Value == null || !firstWorksheet.Cells[counter, 1].Value.ToString().ToLower().Contains("group"))
                {
                    ++counter;
                }
                firstWorksheet.Cells[counter, 2].Value = currGroup.balls[0].GroupNumber;
                // Find line to start entering geometry info
                while (firstWorksheet.Cells[counter, 1].Value == null || !firstWorksheet.Cells[counter, 1].Value.ToString().ToLower().Contains("geometry"))
                {
                    ++counter;
                }
                // Move to first geometry line
                ++counter;
                #region Removed Code
                /*
                // Make some lists to hold the averages of each field for getting the std
                List<double> listHeightAve = new List<double>();
                List<double> listWidthtAve = new List<double>();
                List<double> listTotalAreaAve = new List<double>();
                List<double> listAreaTopAve = new List<double>();
                List<double> listFlatnessAve = new List<double>();
                List<double> listMaxCurvatureAve = new List<double>();
                List<double> listMaxSlopeAve = new List<double>();
                List<double> listMaxSlopeXAve = new List<double>();
                List<double> listMaxSlopeRAve = new List<double>();
                List<double> listSlopeWidthAve = new List<double>();
                List<double> listRecirculationAreaAve = new List<double>();

                List<double> TopHeightAve = new List<double>();
                List<double> TopWidthtAve = new List<double>();
                List<double> TopTotalAreaAve = new List<double>();
                List<double> TopAreaTopAve = new List<double>();
                List<double> TopFlatnessAve = new List<double>();
                List<double> TopMaxCurvatureAve = new List<double>();
                List<double> TopMaxSlopeAve = new List<double>();
                List<double> TopMaxSlopeXAve = new List<double>();
                List<double> TopMaxSlopeRAve = new List<double>();
                List<double> TopSlopeWidthAve = new List<double>();
                List<double> TopRecirculationAreaAve = new List<double>();

                List<double> BottomHeightAve = new List<double>();
                List<double> BottomWidthtAve = new List<double>();
                List<double> BottomTotalAreaAve = new List<double>();
                List<double> BottomAreaTopAve = new List<double>();
                List<double> BottomFlatnessAve = new List<double>();
                List<double> BottomMaxCurvatureAve = new List<double>();
                List<double> BottomMaxSlopeAve = new List<double>();
                List<double> BottomMaxSlopeXAve = new List<double>();
                List<double> BottomMaxSlopeRAve = new List<double>();
                List<double> BottomSlopeWidthAve = new List<double>();
                List<double> BottomRecirculationAreaAve = new List<double>();
                */
                #endregion

                // Save values to their cells
                for (int index = 0; index < 30; ++index)
                {
                    string currGeo = firstWorksheet.Cells[counter, 1].Value.ToString();
                    double[] geoFieldsAverages = currGroup.AveGeometryFields(currGeo);

                    #region Removed code
                    /*
                    // For all 30 tubes
                    listHeightAve.Add(geoFieldsAverages[0]);
                    listWidthtAve.Add(geoFieldsAverages[1]);
                    listTotalAreaAve.Add(geoFieldsAverages[2]);
                    listAreaTopAve.Add(geoFieldsAverages[3]);
                    listFlatnessAve.Add(geoFieldsAverages[4]);
                    listMaxCurvatureAve.Add(geoFieldsAverages[5]);
                    listMaxSlopeAve.Add(geoFieldsAverages[6]);
                    listMaxSlopeXAve.Add(geoFieldsAverages[7]);
                    listMaxSlopeRAve.Add(geoFieldsAverages[8]);
                    listSlopeWidthAve.Add(geoFieldsAverages[9]);
                    listRecirculationAreaAve.Add(geoFieldsAverages[10]);

                    if (index < 15)
                    {
                        // For top tubes
                        TopHeightAve.Add(geoFieldsAverages[0]);
                        TopWidthtAve.Add(geoFieldsAverages[1]);
                        TopTotalAreaAve.Add(geoFieldsAverages[2]);
                        TopAreaTopAve.Add(geoFieldsAverages[3]);
                        TopFlatnessAve.Add(geoFieldsAverages[4]);
                        TopMaxCurvatureAve.Add(geoFieldsAverages[5]);
                        TopMaxSlopeAve.Add(geoFieldsAverages[6]);
                        TopMaxSlopeXAve.Add(geoFieldsAverages[7]);
                        TopMaxSlopeRAve.Add(geoFieldsAverages[8]);
                        TopSlopeWidthAve.Add(geoFieldsAverages[9]);
                        TopRecirculationAreaAve.Add(geoFieldsAverages[10]);
                    }
                    else
                    {
                        // For bottom tubes
                        BottomHeightAve.Add(geoFieldsAverages[0]);
                        BottomWidthtAve.Add(geoFieldsAverages[1]);
                        BottomTotalAreaAve.Add(geoFieldsAverages[2]);
                        BottomAreaTopAve.Add(geoFieldsAverages[3]);
                        BottomFlatnessAve.Add(geoFieldsAverages[4]);
                        BottomMaxCurvatureAve.Add(geoFieldsAverages[5]);
                        BottomMaxSlopeAve.Add(geoFieldsAverages[6]);
                        BottomMaxSlopeXAve.Add(geoFieldsAverages[7]);
                        BottomMaxSlopeRAve.Add(geoFieldsAverages[8]);
                        BottomSlopeWidthAve.Add(geoFieldsAverages[9]);
                        BottomRecirculationAreaAve.Add(geoFieldsAverages[10]);
                    }
                    */
                    #endregion

                    for (int i = 0, column = 6; i < geoFieldsAverages.Length; ++i, column += 2)
                    {
                        firstWorksheet.Cells[counter, column].Value = geoFieldsAverages[i];

                        #region Removed Code
                        /*ave30Tubes[i] += geoFieldsAverages[i];
                        if (firstWorksheet.Cells[counter, 3].Value.ToString().ToLower() == "top")
                        {
                            aveTopTubes[i] += geoFieldsAverages[i];
                        }
                        else if (firstWorksheet.Cells[counter, 3].Value.ToString().ToLower() == "bottom")
                        {
                            aveBottomTubes[i] += geoFieldsAverages[i];
                        }*/
                        #endregion
                    }
                    ++counter;
                }

                #region Removed Code
                /*
                // get averages for 30 tubes
                for (int i = 0; i < ave30Tubes.Length; ++i)
                {
                    ave30Tubes[i] = ave30Tubes[i] / 30.0;
                }

                // get averages for 30 tubes
                for (int i = 0; i < aveTopTubes.Length; ++i)
                {
                    aveTopTubes[i] = aveTopTubes[i] / 15.0;
                }

                // get averages for 30 tubes
                for (int i = 0; i < aveBottomTubes.Length; ++i)
                {
                    aveBottomTubes[i] = aveBottomTubes[i] / 15.0;
                }

                // Get standard deviations of each set
                double[] allStdAverages = new double[11];
                double[] TopStdAverages = new double[11];
                double[] BottomStdAverages = new double[11];
                
                allStdAverages[0] = CalculateStandardDeviation(listHeightAve);
                allStdAverages[1] = CalculateStandardDeviation(listWidthtAve);
                allStdAverages[2] = CalculateStandardDeviation(listTotalAreaAve);
                allStdAverages[3] = CalculateStandardDeviation(listAreaTopAve);
                allStdAverages[4] = CalculateStandardDeviation(listFlatnessAve);
                allStdAverages[5] = CalculateStandardDeviation(listMaxCurvatureAve);
                allStdAverages[6] = CalculateStandardDeviation(listMaxSlopeAve);
                allStdAverages[7] = CalculateStandardDeviation(listMaxSlopeXAve);
                allStdAverages[8] = CalculateStandardDeviation(listMaxSlopeRAve);
                allStdAverages[9] = CalculateStandardDeviation(listSlopeWidthAve);
                allStdAverages[10] = CalculateStandardDeviation(listRecirculationAreaAve);

                TopStdAverages[0] = CalculateStandardDeviation(TopHeightAve);
                TopStdAverages[1] = CalculateStandardDeviation(TopWidthtAve);
                TopStdAverages[2] = CalculateStandardDeviation(TopTotalAreaAve);
                TopStdAverages[3] = CalculateStandardDeviation(TopAreaTopAve);
                TopStdAverages[4] = CalculateStandardDeviation(TopFlatnessAve);
                TopStdAverages[5] = CalculateStandardDeviation(TopMaxCurvatureAve);
                TopStdAverages[6] = CalculateStandardDeviation(TopMaxSlopeAve);
                TopStdAverages[7] = CalculateStandardDeviation(TopMaxSlopeXAve);
                TopStdAverages[8] = CalculateStandardDeviation(TopMaxSlopeRAve);
                TopStdAverages[9] = CalculateStandardDeviation(TopSlopeWidthAve);
                TopStdAverages[10] = CalculateStandardDeviation(TopRecirculationAreaAve);

                BottomStdAverages[0] = CalculateStandardDeviation(BottomHeightAve);
                BottomStdAverages[1] = CalculateStandardDeviation(BottomWidthtAve);
                BottomStdAverages[2] = CalculateStandardDeviation(BottomTotalAreaAve);
                BottomStdAverages[3] = CalculateStandardDeviation(BottomAreaTopAve);
                BottomStdAverages[4] = CalculateStandardDeviation(BottomFlatnessAve);
                BottomStdAverages[5] = CalculateStandardDeviation(BottomMaxCurvatureAve);
                BottomStdAverages[6] = CalculateStandardDeviation(BottomMaxSlopeAve);
                BottomStdAverages[7] = CalculateStandardDeviation(BottomMaxSlopeXAve);
                BottomStdAverages[8] = CalculateStandardDeviation(BottomMaxSlopeRAve);
                BottomStdAverages[9] = CalculateStandardDeviation(BottomSlopeWidthAve);
                BottomStdAverages[10] = CalculateStandardDeviation(BottomRecirculationAreaAve);

                // Insert averages and std
                for (int i = 0, col = 3; i < 11; ++i, col += 2)
                {
                    // 30 Tubes Mean then std
                    firstWorksheet.Cells[11, col].Value = ave30Tubes[i]; 
                    firstWorksheet.Cells[12, col].Value = allStdAverages[i];
                    // Top Tubes Mean then std
                    firstWorksheet.Cells[13, col].Value = aveTopTubes[i]; 
                    firstWorksheet.Cells[14, col].Value = TopStdAverages[i]; 
                    // Bottom tubes Mean then std
                    firstWorksheet.Cells[15, col].Value = aveBottomTubes[i]; 
                    firstWorksheet.Cells[16, col].Value = BottomStdAverages[i];
                }
                */
                #endregion
                // Save the changes
                try
                {
                    excelDoc.Save();
                }
                catch (Exception ex)
                {
                    _log.Error($"Exception when saving excel document: {ex.Message.ToString()}");
                }
            }
        }

        #region STD Calculator
        private static double CalculateStandardDeviation(IEnumerable<double> values)
        {
            double standardDeviation = 0;

            if (values.Any())
            {
                // Compute the average.     
                double avg = values.Average();

                // Perform the Sum of (value-avg)_2_2.      
                double sum = values.Sum(d => Math.Pow(d - avg, 2));

                // Put it all together.      
                standardDeviation = Math.Sqrt((sum) / (values.Count() - 1));
            }

            return standardDeviation;
        }
        #endregion

        private static void ReadData(FileInfo[] files)
        {
            // Each file will be a separate ball
            foreach (FileInfo fi in files)
            {
                if (fi.Name.Split('_').Length < 1)
                {
                    throw new Exception($"Incorrect file name format: {fi.Name}");
                }

                int currLine = 1;
                string ballNum = fi.Name.Split('_')[1];
                Ball currBall = new Ball(fi.Name.Split('_')[0], ballNum);

                // Create a new Excel package from the file
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage excelDoc = new ExcelPackage(fi))
                {
                    try
                    {
                        //Get first WorkSheet. Note that EPPlus indexes start at 1!
                        ExcelWorksheet firstWorksheet = excelDoc.Workbook.Worksheets.FirstOrDefault();
                        while (firstWorksheet.Cells[currLine, 1].Value == null || !firstWorksheet.Cells[currLine, 1].Value.ToString().ToLower().Contains("geometry"))
                        {
                            ++currLine;
                        }
                        _log.Info($"Geometry #s start at worksheet row: {++currLine}, for file name: {fi.Name}");

                        #region Fill each geometry object in a ball report
                        while (firstWorksheet.Cells[currLine, 1].Value != null)
                        {
                            GeometryData gd = new GeometryData();
                            // Geometry #
                            gd.geoNumber = firstWorksheet.Cells[currLine, 1].Value.ToString();
                            gd.geoNumberIco = firstWorksheet.Cells[currLine, 2].Value.ToString();
                            gd.hemisphere = firstWorksheet.Cells[currLine, 3].Value.ToString();
                            gd.geoType = firstWorksheet.Cells[currLine, 5].Value.ToString();

                            // Try to parse an integer values from Icosahedron field
                            if (int.TryParse(firstWorksheet.Cells[currLine, 4].Value.ToString(), out int rslt))
                            {
                                gd.icosahedron = rslt;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Icosahedron from sheet for ball number: {ballNum}");
                            }

                            #region Try parse the rest of the fields as double values
                            // Height
                            if (double.TryParse(firstWorksheet.Cells[currLine, 6].Value.ToString(), out double height))
                            {
                                gd.height = height;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Height from sheet for ball number: {ballNum}");
                            }
                            // Width
                            if (double.TryParse(firstWorksheet.Cells[currLine, 7].Value.ToString(), out double width))
                            {
                                gd.width = width;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Width from sheet for ball number: {ballNum}");
                            }
                            // Total Area
                            if (double.TryParse(firstWorksheet.Cells[currLine, 8].Value.ToString(), out double totalArea))
                            {
                                gd.totalArea = totalArea;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Total Area from sheet for ball number: {ballNum}");
                            }
                            // Area Top
                            if (double.TryParse(firstWorksheet.Cells[currLine, 9].Value.ToString(), out double areaTop))
                            {
                                gd.areaTop = areaTop;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Area Top from sheet for ball number: {ballNum}");
                            }
                            // Flatness
                            if (double.TryParse(firstWorksheet.Cells[currLine, 10].Value.ToString(), out double flatness))
                            {
                                gd.flatness = flatness;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Flatness from sheet for ball number: {ballNum}");
                            }
                            // Max curvature
                            if (double.TryParse(firstWorksheet.Cells[currLine, 11].Value.ToString(), out double curvature))
                            {
                                gd.maxCurvature = curvature;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Max curvature from sheet for ball number: {ballNum}");
                            }
                            // Max Slope Average
                            if (double.TryParse(firstWorksheet.Cells[currLine, 12].Value.ToString(), out double slopeAve))
                            {
                                gd.maxSlopeAve = slopeAve;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Max Slope Average from sheet for ball number: {ballNum}");
                            }
                            // Max Slope X Average
                            if (double.TryParse(firstWorksheet.Cells[currLine, 13].Value.ToString(), out double slopeXAve))
                            {
                                gd.maxSlopeXAve = slopeXAve;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Max Slope X Average from sheet for ball number: {ballNum}");
                            }
                            // Max Slope R Average
                            if (double.TryParse(firstWorksheet.Cells[currLine, 14].Value.ToString(), out double slopeRAve))
                            {
                                gd.maxSlopeRAve = slopeRAve;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Max Slope R Average from sheet for ball number: {ballNum}");
                            }
                            // Slope Width
                            if (double.TryParse(firstWorksheet.Cells[currLine, 15].Value.ToString(), out double slopeWidth))
                            {
                                gd.slopeWidth = slopeWidth;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Slope Width from sheet for ball number: {ballNum}");
                            }
                            // Recirculation area average
                            if (double.TryParse(firstWorksheet.Cells[currLine, 16].Value.ToString(), out double recirc))
                            {
                                gd.recirculationAreaAve = recirc;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Recirculation area average from sheet for ball number: {ballNum}");
                            }
                            #endregion

                            // Add geometry to current ball
                            currBall.AddGeometry(gd);
                            ++currLine;
                        }
                        #endregion
                    }
                    catch (Exception ex)
                    {
                        _log.Error($"Exception in ReadData(): {ex.Message.ToString()}");
                    }
                }
                // Add ball after filling data fields
                currGroup.Add(currBall);
            }
        }
    }
}
