using log4net;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace ExcelWorkerCalla
{
    class MainEntry
    {
        private static readonly ILog _log = LogManager.GetLogger("ExcelWorkerCalla.log");
        private static List<string> _groupNumbers = new List<string>();
        private static List<Group> _allGroups = new List<Group>();
        private static List<string> _controls = new List<string>(); 
        private static string _baseLine = "";

        // Info for use from filename
        private static string[] _info;
        // Given a path to a list of ball excel files (param 1)
        private static string _inputFilePath = "";
        // Ini file path
        private static string _iniFilePath;
        // Given work order number (param 2)
        private static string _workOrderNumber = "";
        // Path to the workorder excel file template
        private static string _excelReportPath;
        // Path to controls template
        private static string _controlsTemplate;
        // Finish stage
        private static string _finishStage;
        static void Main(string[] args)
        {
            _log.Info("Starting executable...");
            for (int i = 0; i < args.Length; ++i)
            {
                _log.Info($"Param #{i + 1}: {args[i]}");
            }

            if (args.Length < 1)
            {
                _log.Error($"{args.Length} arguments passed in, expecting a parameter");
                return;
            }

            if (args.Length != 0)
            {
                _iniFilePath = args[0];
            }

            if (!File.Exists(_iniFilePath))
            {
                _log.Error($"Path to ini file incorrect, {_iniFilePath}");
                return;
            }

            string[] lines = File.ReadAllLines(_iniFilePath);
            foreach (string line in lines)
            {
                if (line.Split('=')[0] == "inputScanData")
                {
                    _inputFilePath = line.Split('=')[1];
                }
                if (line.Split('=')[0] == "workOrder")
                {
                    _workOrderNumber = line.Split('=')[1];
                }
            }
            
            if (_workOrderNumber == "")
            {
                _log.Error("Work order number not found in ini file..");
                return;
            }

            // Get directories needed
            string tempFolder = Path.GetDirectoryName(_inputFilePath);
            string golfBallFolder = Path.GetDirectoryName(tempFolder);
            string reportsFolder = golfBallFolder + "\\Reports";
            string archiveFolder = golfBallFolder + "\\Archive";
            _excelReportPath = archiveFolder + "\\WorkOrder Report Template.xlsx";
            _controlsTemplate = archiveFolder + "\\Controls Template.xlsx";

            // Validate both directories
            if (!File.Exists(_inputFilePath))
            {
                _log.Error($"Invalid directory {_inputFilePath}");
                return;
            }

            if (!Directory.Exists(golfBallFolder))
            {
                _log.Error($"Invalid path {golfBallFolder}");
                return;
            }

            // Check if directory exists
            if (!Directory.Exists(reportsFolder))
            {
                _log.Error($"Directory: {reportsFolder} does not exist.");
                return;
            }
            
            // Get Work Order reports path
            string workOrderFolder = reportsFolder + "\\02_WorkOrder Reports";

            // Check if the work order folder exists
            if (!Directory.Exists(workOrderFolder))
            {
                _log.Error($"Directory: {workOrderFolder} does not exist...");
                return;
            }

            // Copy template to workOrderFolder and rename it
            // Get aspects of file name, 0 = date, 1 = time
            _info = Path.GetFileName(_inputFilePath).Split('_');
            string newReportFilePath = workOrderFolder + "\\" + _info[0] + "_" + _info[1] + "_" + _workOrderNumber + "_Report.xlsx";
            try
            {
                File.Copy(_excelReportPath, newReportFilePath);
            }
            catch (Exception ex)
            {
                _log.Error($"Error when copying work order template file to new location: {newReportFilePath}, error message: {ex.Message.ToString()}");
                return;
            }

            // Create a new Excel package from the file
            FileInfo inputFiles = new FileInfo(_inputFilePath);

            // Parse input file for needed info like potential group numbers, controls _controls and the baseline _baseLine used
            ParseInputFile(inputFiles);
            // Now we have a list of group numbers _groupNumbers, controls and a baseline

            // Need to search the report folder for all related group sub folders with .xlsx data files to read from
            List<string> groupFolderPaths = new List<string>();

            // Make a list of folders to parse 
            string[] allFolders = Directory.GetDirectories(reportsFolder);

            // If a folder name includes a group number we want then parse it
            foreach (string folderPath in allFolders)
            {
                string folder = Path.GetFileName(folderPath);
                string date = folder.Split('_')[0];
                if (date != _info[0])
                {
                    continue;
                }

                string time = folder.Split('_')[1];
                if (time != _info[1])
                {
                    continue;
                }

                string grpNum = folder.Split('_')[2];
                if (grpNum != null && _groupNumbers.Contains(grpNum))
                {
                    groupFolderPaths.Add(folderPath);
                }
            }

            // Now we have paths to each group folder that we need to record
            // Should have equal number of folder paths and group numbers listed, so check
            if (groupFolderPaths.Count != _groupNumbers.Count)
            {
                _log.Error($"Number of group numbers in folder hierarchy({groupFolderPaths.Count}) does not match quantity found in input file({_groupNumbers.Count}) {Environment.NewLine}for Work Order number: {_workOrderNumber}..");
            }

            // Loop through all relevant folders full of ball data
            foreach (string groupFolder in groupFolderPaths)
            {
                string currGrpNumber = Path.GetFileName(groupFolder).Split('_')[2];
                try
                {
                    DirectoryInfo direct = new DirectoryInfo(groupFolder);
                    FileInfo[] files = direct.GetFiles("*.xlsx");
                    _allGroups.Add(ReadData(files, currGrpNumber));
                }
                catch (Exception ex)
                {
                    _log.Error($"Exception thrown when looping through ball files and directories: {ex.Message.ToString()}");
                    return;
                }
            }
            
            // Create a new Excel package from the copied excel file
            FileInfo fi = new FileInfo(newReportFilePath);

            // Open excel template file
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excelDoc = new ExcelPackage(fi))
            {
                // Copy Baseline and Control sheets from template to this work order
                using (ExcelPackage excelTemp = new ExcelPackage(new FileInfo(_controlsTemplate)))
                {
                    // Check if the template worksheet has a baseline sheet, if not return
                    if (excelTemp.Workbook.Worksheets.FirstOrDefault(x => x.Name == "Baseline") == null)
                    {
                        _log.Error($"Baseline worksheet does not exist in the controls template..");
                        return;
                    }
                    // Add the baseline to the new workorder doc
                    excelDoc.Workbook.Worksheets.Add("Baseline", excelTemp.Workbook.Worksheets["Baseline"]);
                    // Add all the control sheets from template to new work order doc
                    foreach (string sheetName in _controls)
                    {
                        if (excelTemp.Workbook.Worksheets.FirstOrDefault(x => x.Name == sheetName) == null)
                        {
                            _log.Error($"Worksheet {sheetName} does not exist in the controls template..");
                            return;
                        }
                        else
                        {
                            excelDoc.Workbook.Worksheets.Add(sheetName, excelTemp.Workbook.Worksheets[sheetName]);
                        }
                    }
                }

                // Create new sheet named : group inside excel report for each group in _groupNumbers
                foreach (string grp in _groupNumbers)
                {
                    excelDoc.Workbook.Worksheets.Copy("Group", grp);
                }

                // Remove generic group sheet
                excelDoc.Workbook.Worksheets.Delete("Group");

                // Loop through all groups and fill the data accordingly
                for (int i = 0; i < _groupNumbers.Count; ++i)
                {
                    foreach (Group _currGroup in _allGroups)
                    {
                        if (_currGroup.GroupNumber == _groupNumbers[i])
                        {
                            // open sheet with group number as the name
                            //Get worksheet of that group number and write to it
                            ExcelWorksheet currSheet = excelDoc.Workbook.Worksheets[_groupNumbers[i]];

                            // Find line to insert group number on
                            int counter = 1;
                            while (currSheet.Cells[counter, 1].Value == null || !currSheet.Cells[counter, 1].Value.ToString().ToLower().Contains("group"))
                            {
                                ++counter;
                            }

                            // Write Info atop each group sheet
                            currSheet.Cells[1, 1].Value = $"Ball Scan Report \n WO {_workOrderNumber}({_baseLine} - Baseline)";

                            // Write group number on top of sheet
                            currSheet.Cells[counter, 2].Value = _groupNumbers[i];

                            // Write baseline/CAD number to top sheet
                            currSheet.Cells[counter + 2, 2].Value = _baseLine;

                            // Find line to start entering geometry info
                            while (currSheet.Cells[counter, 1].Value == null || !currSheet.Cells[counter, 1].Value.ToString().ToLower().Contains("geometry"))
                            {
                                ++counter;
                            }

                            // Move to first geometry line
                            ++counter;

                            // Save values to their cells
                            for (int index = 0; index < 30; ++index, ++counter)
                            {
                                string currGeo = currSheet.Cells[counter, 1].Value.ToString();
                                double[] geoFieldsAverages = _currGroup.AveGeometryFields(currGeo);


                                for (int ind = 0, column = 6; ind < geoFieldsAverages.Length; ++ind, column += 2)
                                {
                                    currSheet.Cells[counter, column].Value = geoFieldsAverages[ind];

                                }
                            }
                        }
                    }
                }

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

        private static void ParseInputFile(FileInfo file)
        {
            // Create a new Excel package from the file
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excelDoc = new ExcelPackage(file))
            {
                // _groupNumbers
                // Get first worksheet to grab data from
                ExcelWorksheet rackOneAndTwo = excelDoc.Workbook.Worksheets.FirstOrDefault();

                // Rack 1
                AnalyzeSection(rackOneAndTwo, 3, 1, 10, 3);
                AnalyzeSection(rackOneAndTwo, 3, 7, 10, 9);
                AnalyzeSection(rackOneAndTwo, 3, 13, 10, 15);
                AnalyzeSection(rackOneAndTwo, 3, 19, 10, 21);
                AnalyzeSection(rackOneAndTwo, 3, 25, 10, 27);
                AnalyzeSection(rackOneAndTwo, 3, 31, 10, 33);
                AnalyzeSection(rackOneAndTwo, 3, 37, 10, 39);
                AnalyzeSection(rackOneAndTwo, 3, 43, 10, 45);
                AnalyzeSection(rackOneAndTwo, 3, 49, 10, 51);
                AnalyzeSection(rackOneAndTwo, 3, 55, 10, 57);

                // Rack 2
                AnalyzeSection(rackOneAndTwo, 26, 1, 33, 3);
                AnalyzeSection(rackOneAndTwo, 26, 7, 33, 9);
                AnalyzeSection(rackOneAndTwo, 26, 13, 33, 15);
                AnalyzeSection(rackOneAndTwo, 26, 19, 33, 21);
                AnalyzeSection(rackOneAndTwo, 26, 25, 33, 27);
                AnalyzeSection(rackOneAndTwo, 26, 31, 33, 33);
                AnalyzeSection(rackOneAndTwo, 26, 37, 33, 39);
                AnalyzeSection(rackOneAndTwo, 26, 43, 33, 45);
                AnalyzeSection(rackOneAndTwo, 26, 49, 33, 51);
                AnalyzeSection(rackOneAndTwo, 26, 55, 33, 57);

            }
        }

        /* Need: 
                 * 1 - WO# Top start position (Row/Col) ex. 3,1
                 * 2 - Group# start (Row/Col) ex. 10,3 
        */
        private static void AnalyzeSection(ExcelWorksheet rackOneAndTwo, int wOTopStartRow, int wOTopStartCol, int groupStartRow, int groupStartCol)
        {
            // Rack 1, Row 1, designed  to loop through the 4 possible rows with WO#, baseline/CAD, and controls
            for (int i = 0; i < 4; ++i)
            {
                // Check WO #s at top each row, maximum of 4 and only read if there is a value and it matches the current work order
                if (rackOneAndTwo.Cells[i + wOTopStartRow, wOTopStartCol].Value != null && rackOneAndTwo.Cells[i + wOTopStartRow, wOTopStartCol].Value.ToString() == _workOrderNumber)
                {
                    // Add CAD and controls listed for that WO number
                    if (_baseLine == "")
                    {
                        // Baseline is always first before the controls
                        _baseLine = rackOneAndTwo.Cells[i + wOTopStartRow, wOTopStartCol + 1].Value.ToString();
                    }

                    // Add controls if they are not added yet, up to 3 controls per row
                    for (int count = wOTopStartCol + 2; count < wOTopStartCol + 5; ++count)
                    {
                        // Check each colum for the remainder of the controls if value is not null
                        if (rackOneAndTwo.Cells[i + wOTopStartRow, count].Value != null && !_controls.Contains(rackOneAndTwo.Cells[i + wOTopStartRow, count].Value.ToString()))
                        {
                            // If does not exist in the list then add it
                            _controls.Add(rackOneAndTwo.Cells[i + wOTopStartRow, count].Value.ToString());
                        }
                    }

                    // Loop through cells to grab group numbers involved, max number of 10 per row
                    for (int j = 0; j < 10; ++j)
                    {
                        // Check the WO#, if it matches then add the group number in the column before it
                        if (rackOneAndTwo.Cells[i + groupStartRow, groupStartCol + 1].Value != null && rackOneAndTwo.Cells[i + groupStartRow, groupStartCol + 1].Value.ToString() == _workOrderNumber)
                        {
                            if (rackOneAndTwo.Cells[i + groupStartRow, groupStartCol].Value == null)
                            {
                                _log.Error($"No Value found in input sheet at position {i + groupStartRow}, {groupStartCol} where there is a value for WO# in the next cell.");
                                return;
                            }
                            // First part of split holds group number, second is ball number
                            string grpNum = rackOneAndTwo.Cells[i + groupStartRow, groupStartCol].Value.ToString().Split('_')[0];

                            // Add it if it is not already in the list
                            if (!_groupNumbers.Contains(grpNum))
                            {
                                _groupNumbers.Add(grpNum);
                            }
                        }
                    }
                }
            }
        }

        // Read the files passed in to compile the group data in report sheet
        private static Group ReadData(FileInfo[] files, string grpnum)
        {
            Group currGroup = new Group(grpnum);
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
                            if (firstWorksheet.Cells[currLine, 4].Value != null && int.TryParse(firstWorksheet.Cells[currLine, 4].Value.ToString(), out int rslt))
                            {
                                gd.icosahedron = rslt;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Icosahedron from sheet for ball number: {ballNum}");
                            }

                            #region Try parse the rest of the fields as double values
                            // Height
                            if (firstWorksheet.Cells[currLine, 6].Value != null && double.TryParse(firstWorksheet.Cells[currLine, 6].Value.ToString(), out double height))
                            {
                                gd.height = height;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Height from sheet for ball number: {ballNum}");
                            }
                            // Width
                            if (firstWorksheet.Cells[currLine, 7].Value != null && double.TryParse(firstWorksheet.Cells[currLine, 7].Value.ToString(), out double width))
                            {
                                gd.width = width;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Width from sheet for ball number: {ballNum}");
                            }
                            // Total Area
                            if (firstWorksheet.Cells[currLine, 8].Value != null && double.TryParse(firstWorksheet.Cells[currLine, 8].Value.ToString(), out double totalArea))
                            {
                                gd.totalArea = totalArea;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Total Area from sheet for ball number: {ballNum}");
                            }
                            // Area Top
                            if (firstWorksheet.Cells[currLine, 9].Value != null && double.TryParse(firstWorksheet.Cells[currLine, 9].Value.ToString(), out double areaTop))
                            {
                                gd.areaTop = areaTop;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Area Top from sheet for ball number: {ballNum}");
                            }
                            // Flatness
                            if (firstWorksheet.Cells[currLine, 10].Value != null && double.TryParse(firstWorksheet.Cells[currLine, 10].Value.ToString(), out double flatness))
                            {
                                gd.flatness = flatness;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Flatness from sheet for ball number: {ballNum}");
                            }
                            // Max curvature
                            if (firstWorksheet.Cells[currLine, 11].Value != null && double.TryParse(firstWorksheet.Cells[currLine, 11].Value.ToString(), out double curvature))
                            {
                                gd.maxCurvature = curvature;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Max curvature from sheet for ball number: {ballNum}");
                            }
                            // Max Slope Average
                            if (firstWorksheet.Cells[currLine, 12].Value != null && double.TryParse(firstWorksheet.Cells[currLine, 12].Value.ToString(), out double slopeAve))
                            {
                                gd.maxSlopeAve = slopeAve;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Max Slope Average from sheet for ball number: {ballNum}");
                            }
                            // Max Slope X Average
                            if (firstWorksheet.Cells[currLine, 13].Value != null && double.TryParse(firstWorksheet.Cells[currLine, 13].Value.ToString(), out double slopeXAve))
                            {
                                gd.maxSlopeXAve = slopeXAve;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Max Slope X Average from sheet for ball number: {ballNum}");
                            }
                            // Max Slope R Average
                            if (firstWorksheet.Cells[currLine, 14].Value != null && double.TryParse(firstWorksheet.Cells[currLine, 14].Value.ToString(), out double slopeRAve))
                            {
                                gd.maxSlopeRAve = slopeRAve;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Max Slope R Average from sheet for ball number: {ballNum}");
                            }
                            // Slope Width
                            if (firstWorksheet.Cells[currLine, 15].Value != null && double.TryParse(firstWorksheet.Cells[currLine, 15].Value.ToString(), out double slopeWidth))
                            {
                                gd.slopeWidth = slopeWidth;
                            }
                            else
                            {
                                _log.Error($"Failed to parse Slope Width from sheet for ball number: {ballNum}");
                            }
                            // Recirculation area average
                            if (firstWorksheet.Cells[currLine, 16].Value != null && double.TryParse(firstWorksheet.Cells[currLine, 16].Value.ToString(), out double recirc))
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
            return currGroup;
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

    }
}
