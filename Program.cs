#region Using statements
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;

using databaseAPI;

using GNA_CommercialLicenseValidator;

using GNAchartingtools;

using gnaDataClasses;

using GNAgeneraltools;

using GNAspreadsheettools;

using GNAsurveytools;

using OfficeOpenXml;

using T4Dlibrary;

using Twilio.TwiML.Voice;
using Twilio.Types;

using static GNAgeneraltools.gnaTools;

#endregion

namespace GNA_StructuralDisplacementReport
{
    #region Internal helper methods
    internal static class ConfigParsing
    {
        public static bool GetBoolYesNo(System.Collections.Specialized.NameValueCollection appSettings, string key)
        {
            string? raw = appSettings[key];
            if (raw is null)
                throw new ConfigurationErrorsException($"Missing required appSetting: '{key}'.");

            string value = raw.Trim();

            if (value.Equals("Yes", StringComparison.OrdinalIgnoreCase)) return true;
            if (value.Equals("No", StringComparison.OrdinalIgnoreCase)) return false;

            throw new ConfigurationErrorsException(
                $"Invalid value for appSetting '{key}': '{raw}'. Expected 'Yes' or 'No'.");
        }

        public static string GetRequiredString(System.Collections.Specialized.NameValueCollection appSettings, string key)
        {
            string? raw = appSettings[key];
            if (string.IsNullOrWhiteSpace(raw))
                throw new ConfigurationErrorsException($"Missing or empty required appSetting: '{key}'.");
            return raw.Trim();
        }

        public static int GetRequiredInt(System.Collections.Specialized.NameValueCollection appSettings, string key)
        {
            string raw = GetRequiredString(appSettings, key);
            if (!int.TryParse(raw, out int value))
                throw new ConfigurationErrorsException($"Invalid integer for appSetting '{key}': '{raw}'.");
            return value;
        }

        public static double GetRequiredDouble(System.Collections.Specialized.NameValueCollection appSettings, string key)
        {
            string raw = GetRequiredString(appSettings, key);
            if (!double.TryParse(
                    raw,
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out double value))
            {
                throw new ConfigurationErrorsException(
                    $"Invalid double for appSetting '{key}': '{raw}'. Use '.' as decimal separator.");
            }
            return value;
        }
    }


    #endregion

    internal class Program
    {
        static void Main()
        {
            try
            {
#pragma warning disable CS0162
#pragma warning disable CS8600
#pragma warning disable CS8601
#pragma warning disable CS8602
#pragma warning disable CS8604
#pragma warning disable IDE0028
#pragma warning disable IDE0059

                //===============[Initial settings]======================================

                #region Initial setup
                gnaTools gnaT = new();
                GNAsurveycalcs gnaSurvey = new();
                dbAPI gnaDBAPI = new();
                spreadsheetAPI gnaSpreadsheetAPI = new spreadsheetAPI(db: gnaDBAPI);
                GNAchartingAPI chartingAPI = new();
                T4Dapi t4dapi = new();

                string strTab1 = "     ";
                string strTab2 = "        ";

                Console.OutputEncoding = System.Text.Encoding.Unicode;
                Console.Clear();

                gnaT.WelcomeMessage($"GNA_StructuralDisplacementReport {BuildInfo.BuildDateString()}");
                #endregion

                #region Check Config file and license
                int headingNo = 1;
                Console.WriteLine($"{headingNo++}. Checking license and config file");
                Console.WriteLine($"{strTab1}Verify config file");
                gnaT.VerifyLocalConfig();

                var config = ConfigurationManager.AppSettings;

                Console.WriteLine($"{strTab1}Validating software licenses");
                string licenseCode = config["LicenseCode"] ?? string.Empty;
                if (string.IsNullOrWhiteSpace(licenseCode))
                {
                    Console.WriteLine($"{strTab2}License code is not set in the configuration file.");
                    return;
                }

                LicenseValidator.ValidateLicense("DSPRPT", licenseCode);
                Console.WriteLine($"{strTab2}Validated");

                // Set the T4DAPI license
                string T4DAPI_licenseCode = config["T4Dapi_LicenseCode"] ?? string.Empty;
                t4dapi.SetCommercial(T4DAPI_licenseCode);
                gnaT.epplusLicense();
                Console.WriteLine($"{strTab1}Done");
                #endregion

                #region Config Variables
                Console.WriteLine($"{headingNo++}. System variables");

                string strDBconnection = ConfigurationManager.ConnectionStrings["DBconnectionString"].ConnectionString;

                bool freezeScreen = ConfigParsing.GetBoolYesNo(config, "freezeScreen");
                bool computedRdT = ConfigParsing.GetBoolYesNo(config, "computedRdT");
                bool ATScallibration = ConfigParsing.GetBoolYesNo(config, "computeATScallibration");
                bool manualSurvey = ConfigParsing.GetBoolYesNo(config, "ManualSurvey");
                bool drawCharts = ConfigParsing.GetBoolYesNo(config, "DrawCharts");
                bool debug = ConfigParsing.GetBoolYesNo(config, "Debug");   

                string strFreezeScreen = freezeScreen ? "Yes" : "No";

                string strProjectTitle = ConfigParsing.GetRequiredString(config, "ProjectTitle");


                string strClient = ConfigParsing.GetRequiredString(config, "Client");
                int iFirstDataRow = ConfigParsing.GetRequiredInt(config, "FirstDataRow");
                int iFirstDataCol = ConfigParsing.GetRequiredInt(config, "FirstDataCol");
                int iFirstOutputRow = ConfigParsing.GetRequiredInt(config, "FirstOutputRow");

                string strReferenceLineTerminals = ConfigParsing.GetRequiredString(config,key: "ReferenceLineTerminalsEaNaEbNb");



                double dblDataJumpTriggerLevel = ConfigParsing.GetRequiredDouble(config, "DataJumpTriggerLevel");

                string strTimeBlockType = ConfigParsing.GetRequiredString(config, "TimeBlockType");
                string strManualBlockStart = ConfigParsing.GetRequiredString(config, "manualBlockStart");
                string strManualBlockEnd = ConfigParsing.GetRequiredString(config, "manualBlockEnd");
                string strBlockSizeHrs = ConfigParsing.GetRequiredString(config, "BlockSizeHrs");

                int iNoOfTimeBlocksPerReport = ConfigParsing.GetRequiredInt(config, "NoOfTimeBlocksPerReport");
                int iNoOfEpochsHistoricData = ConfigParsing.GetRequiredInt(config, "NoOfEpochsHistoricData");
                #endregion

                #region Workbook variables

                // Always required
                string strExcelPath = ConfigParsing.GetRequiredString(config, "ExcelPath");
                string strExcelFile = ConfigParsing.GetRequiredString(config, "ExcelFile");
                string strReferenceWorksheet = ConfigParsing.GetRequiredString(config, "ReferenceWorksheet");

                // Optional (may be absent from config AND from this executable's usage)
                string? strSurveyWorksheet = config["SurveyWorksheet"]?.Trim();
                string? strHistoricCoordinatesWorksheet = config["HistoricCoordinatesWorksheet"]?.Trim();
                string? strHistoricDeltasWorksheet = config["HistoricDeltasWorksheet"]?.Trim();
                string? strHistoricDRWorksheet = config["HistoricDRWorksheet"]?.Trim();
                string? strHistoricDTWorksheet = config["HistoricDTWorksheet"]?.Trim();
                string? strHistoricDHWorksheet = config["HistoricDHWorksheet"]?.Trim();
                string? strHistoricDistanceWorksheet = config["HistoricDistanceWorksheet"]?.Trim();
                string? strLatestCoordinatesWorksheet = config["LatestCoordinatesWorksheet"]?.Trim();
                string? strLatestPolarDisplacementsWorksheet = config["LatestPolarDisplacementsWorksheet"]?.Trim();
                string? strCalibrationWorksheet = config["CalibrationWorksheet"]?.Trim();



                #endregion

                #region Report variables
                string strContractTitle = ConfigParsing.GetRequiredString(config, "ContractTitle");
                string strReportType = ConfigParsing.GetRequiredString(config, "ReportType");
                #endregion

                #region System Variables
                string strMasterWorkbookFullPath = strExcelPath + strExcelFile;
                Console.WriteLine($"{strTab1}Done");
                #endregion

                #region Clean exit
                void FinishAndExit()
                {
                    Console.WriteLine("\nStructural displacement report completed...\n\n");
                    gnaT.freezeScreen(strFreezeScreen);
                }
                #endregion

                #region Check whether workbook is open

                if (gnaSpreadsheetAPI.IsWorkbookOpen(
                    strWorkbookFullPath: strMasterWorkbookFullPath))
                {
                    string message =
                        $"{strTab1}The Excel workbook is currently open or locked:\n " +
                        $"'{strMasterWorkbookFullPath}'.";
                    Console.WriteLine($"{message}\nExecution stopped...\n");
                    Environment.Exit(exitCode: 0);
                }
                else
                {
                    Console.WriteLine($"{strTab1}{strMasterWorkbookFullPath} ready");
                }
                #endregion

                #region populate the RuntimeEnvironment class

                RuntimeEnvironment env = new RuntimeEnvironment
                {
                    // ---- Database ----
                    DbConnectionString = strDBconnection,
                    ProjectTitle = strProjectTitle,

                    // ---- Workbook ----
                    ExcelPath = strExcelPath,
                    ExcelFile = strExcelFile,

                    // ---- Worksheets ----
                    ReferenceWorksheet = strReferenceWorksheet,
                    SurveyWorksheet = strSurveyWorksheet,
                    HistoricCoordinatesWorksheet = strHistoricCoordinatesWorksheet,
                    HistoricDeltasWorksheet = strHistoricDeltasWorksheet,
                    HistoricDRWorksheet = strHistoricDRWorksheet,
                    HistoricDTWorksheet = strHistoricDTWorksheet,
                    HistoricDHWorksheet = strHistoricDHWorksheet,
                    HistoricDistanceWorksheet = strHistoricDistanceWorksheet,   
                    CalibrationWorksheet = strCalibrationWorksheet, 
                    LatestCoordinatesWorksheet = strLatestCoordinatesWorksheet,
                    LatestPolarDisplacementsWorksheet = strLatestPolarDisplacementsWorksheet,

                    // ---- Row/Col configuration ----
                    FirstDataRow = iFirstDataRow,
                    FirstDataCol = iFirstDataCol,
                    FirstOutputRow = iFirstOutputRow
                };

                #endregion

                #region ATS details
                // read the ats and settop from the config file and write to AlarmState
                Console.WriteLine($"{headingNo++}. Extract ATS and settop data");
                List<ATS> atsList = AtsConfigurationReader.extractATSdetails(env: env);

                if (ATScallibration)
                {
                    if (atsList is null || atsList.Count == 0)
                    {
                        throw new InvalidOperationException($"{strTab1}No ATS records were extracted from configuration.");
                    }
                    else
                    {
                        Console.WriteLine($"{strTab1}Extracted {atsList.Count} ATS records");
                    }
                }
                else
                {
                    Console.WriteLine($"{strTab1}No ATS calibration");
                }
                // env already constructed above this point
                #endregion

                #region Environment check
                Console.WriteLine($"{headingNo++}. Check system environment");
                Console.WriteLine($"{strTab1}Project: " + strProjectTitle);
                Console.WriteLine($"{strTab1}Master workbook: " + strMasterWorkbookFullPath);

                if (freezeScreen)
                {
                    Console.WriteLine($"{strTab1}Check database connection");
                    gnaDBAPI.testDBconnection(strDBconnection);

                    Console.WriteLine($"{strTab1}Check Existance of workbook & worksheets");
                    gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strReferenceWorksheet);
                    gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strSurveyWorksheet);

                    // NOTE: The following worksheet variables are referenced in your original code but were not
                    // present in the provided excerpt as defined config keys/variables. They are intentionally
                    // not referenced here to keep this file self-consistent and compilable as provided.

                    if (ATScallibration)
                    {
                        gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strHistoricDistanceWorksheet);
                        gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strCalibrationWorksheet);
                    }
                    else
                    {
                        Console.WriteLine($"{strTab1}Existance of Calibration worksheet not checked");
                    }

                    if (computedRdT)
                    {
                        gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strHistoricDRWorksheet);
                        gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strHistoricDTWorksheet);
                        gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strHistoricDHWorksheet);
                    }
                    else
                    {
                        Console.WriteLine($"{strTab1}Existance of historic dR & dT worksheets not checked");
                    }

                    if (drawCharts)
                    {
                        // Charts worksheet checks belong here when worksheet names are available.
                    }
                    else
                    {
                        Console.WriteLine($"{strTab1}Existance of Charts worksheets not checked");
                    }
                }
                else
                {
                    Console.WriteLine($"{strTab1}Existance of workbook & worksheets is not checked");
                }

                // NOTE: iNoOfPrisms usage removed as unused in provided code.
                Console.WriteLine($"{strTab1}Done");
                #endregion

                #region Time blocks
                Console.WriteLine($"{headingNo++}. Time blocks");
                List<Tuple<string, string>> subBlocks = new List<Tuple<string, string>>();

                double dblTimeZoneOffset = gnaDBAPI.getProjectTimeZoneOffset(strDBconnection, strProjectTitle);
                Console.WriteLine($"{strTab1}Project time zone offset: {dblTimeZoneOffset} hrs");

                switch (strTimeBlockType)
                {
                    case "Historic":
                        subBlocks = gnaT.prepareTimeBlocksWithTimeZoneOffset(
                            strTimeBlockType: "Historic",
                            strBlockSizeHrs: strBlockSizeHrs,
                            strManualBlockStart: strManualBlockStart,
                            strManualBlockEnd: strManualBlockEnd,
                            dblTimeZoneOffset: dblTimeZoneOffset);
                        break;

                    case "Manual":
                        subBlocks = gnaT.prepareTimeBlocksWithTimeZoneOffset(
                            strTimeBlockType: "Manual",
                            strManualBlockStart: strManualBlockStart,
                            strManualBlockEnd: strManualBlockEnd,
                            dblTimeZoneOffset: dblTimeZoneOffset);
                        break;

                    case "Schedule":
                        subBlocks = gnaT.prepareTimeBlocksWithTimeZoneOffset(
                            strTimeBlockType: "Schedule",
                            strBlockSizeHrs: strBlockSizeHrs,
                            dblTimeZoneOffset: dblTimeZoneOffset);
                        break;

                    default:
                        Console.WriteLine("\nError in Timeblock Type");
                        Console.WriteLine("Time block type: " + strTimeBlockType);
                        Console.WriteLine("Must be Manual, Schedule or Historic");
                        Console.WriteLine("\nPress key to exit...");
                        Console.ReadKey();
                        Environment.Exit(1);
                        break;
                }

                Console.WriteLine($"{strTab1}Done");
                #endregion

                #region Survey Worksheet Preparation
                Console.WriteLine($"{headingNo++}. {strSurveyWorksheet} preparation");

                if (!manualSurvey)
                {
                    Console.WriteLine($"{strTab1}Read point names");
                    string[] strPointNames = gnaSpreadsheetAPI.readPointNames(
                        strMasterWorkbookFullPath,
                        strSurveyWorksheet,
                        iFirstDataRow.ToString(System.Globalization.CultureInfo.InvariantCulture));

                    Console.WriteLine($"{strTab1}Extract SensorID");
                    string[,] strSensorID = gnaDBAPI.getSensorIDfromDB(strDBconnection, strPointNames, strProjectTitle);

                    if (debug)
                    {
                        int counter = 0;
                        Console.WriteLine($"\nstrProjectTitle: {strProjectTitle}");

                        while (counter < strSensorID.GetLength(0))
                        {
                            string name = (strSensorID[counter, 0] ?? string.Empty).Trim();
                            if (name == "NoMore") break;
                            string id = (strSensorID[counter, 1] ?? string.Empty).Trim();
                            Console.WriteLine($"{counter}  {name}  {id}");
                            counter++;
                        }
                        Console.WriteLine("\n");
                    }

                    Console.WriteLine($"{strTab1}Update SensorID");
                    gnaSpreadsheetAPI.writeSensorID(
                        strMasterWorkbookFullPath,
                        strSurveyWorksheet,
                        strSensorID,
                        iFirstDataRow.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    Console.WriteLine($"{strTab1}Done");
                }
                else
                {
                    Console.WriteLine($"{strTab1}Manual Survey:\n{strTab2}Jump steps 2,3,4  ({strClient})");
                }
                #endregion


                #region Time block processing
                string strDateTime = DateTime.Now.ToString("yyyyMMdd_HHmm");
                string strDateTimeUTC = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss");   //2022-07-26 13:45:15
                string strTimeStamp = "";
                string strReportTime = "";
                string strExportFile = "";

                Console.WriteLine($"{headingNo++}. Block Processing");
                Console.WriteLine($"{strTab1}Timeblock Type: {strTimeBlockType}");

                foreach (var block in subBlocks)
                {
                    string strTimeBlockStartUTC = block.Item1;
                    string strTimeBlockEndUTC = block.Item2;
                    string strTimeBlockStartLocal = gnaT.convertUTCToLocalWithTimeZoneOffset(strTimeBlockStartUTC, dblTimeZoneOffset).Trim();
                    string strTimeBlockEndLocal = gnaT.convertUTCToLocalWithTimeZoneOffset(strTimeBlockEndUTC, dblTimeZoneOffset).Trim();

                    strReportTime = strTimeBlockEndLocal;
                    strReportTime = strReportTime.Replace("-", "");
                    strReportTime = strReportTime.Replace(" ", "_");
                    strReportTime = strReportTime.Replace(":", "");

                    Console.WriteLine($"{strTab2}{strTimeBlockStartLocal} (local)");
                    Console.WriteLine($"{strTab2}{strTimeBlockEndLocal} (local)");

                    if ((strTimeBlockType == "Manual") || (strTimeBlockType == "Historic") || (manualSurvey))
                    {
                        strExportFile = strExcelPath + strContractTitle + "_" + strReportType + "_" + strReportTime + "m.xlsx";
                    }
                    else
                    {
                        strExportFile = strExcelPath + strContractTitle + "_" + strReportType + "_" + strReportTime + ".xlsx";
                    }

                    strTimeStamp = strTimeBlockEndLocal + "\n(local)";

                    if (manualSurvey)
                    {
                        Console.WriteLine($"{strTab1}Manual survey - no database extraction");
                        strTimeStamp = strTimeBlockEndLocal + "\nManual Survey";
                    }
                    else
                    {
                        string strBlockStart = gnaT.NormalizeTimeStampToString(strTimeBlockStartUTC);
                        string strBlockEnd = gnaT.NormalizeTimeStampToString(strTimeBlockEndUTC);

                        #region Calibration worksheet processing
                        Console.WriteLine($"{strTab1}Calibration worksheet");
                        if (ATScallibration)
                        {
                            Console.WriteLine($"{strTab2}Slope distances");
                            // Retrieve mean slope distances for the time block
                            // Write to Historic Dist and to Calibration worksheets
                            List<CalibrationData> data = gnaSpreadsheetAPI.processMeanSlopeDistance(
                                env: env,
                                strBlockStart: strBlockStart,
                                strBlockEnd: strBlockEnd);

                            if (data is null)
                            {
                                Console.WriteLine($"{strTab1}Error: no data retrieved");
                                throw new InvalidOperationException("retrieveMeanSlopeDistance returned null.");
                            }




                        }
                        else {
                            Console.WriteLine($"{strTab2}No calibration");
                            
                        }

                        #endregion


                        //==============================================================
                        Console.WriteLine($"{strTab1}Extract mean deltas for time block");
                        List<Points> currentCoordinates = t4dapi.GetAllPointsMeanDeltas(strDBconnection, strProjectTitle,
                                    strBlockStart, strBlockEnd);

                        Console.WriteLine($"{strTab2}Count returned: {currentCoordinates.Count}");

                        // Update with reference data and compute current coordinates
                        Console.WriteLine($"{strTab1}Update Workbook");
                        Console.WriteLine($"{strTab2}Write Coordinates");
                        currentCoordinates = gnaSpreadsheetAPI.updateCurrentCoordinates(
                            env: env,
                            currentCoordinates: currentCoordinates,
                            strTimeBlockEndUTC: strTimeBlockEndUTC);

                        // clean current coordinates by removing the null value coordinates
                        currentCoordinates = t4dapi.RemovePointsWithNullEref(currentCoordinates);

                        // echo current coordinates to console for debugging
                        if (debug)
                            {
                            Console.WriteLine("Current coordinates");
                                foreach (var c in currentCoordinates)
                                {
                                    Console.WriteLine(
                                        $"SensorID:{c.SensorID}\n" +
                                        $"Name:{c.Name}\n" +
                                        $"ReplacementName:{c.ReplacementName}\n" +
                                        $"Count:{c.Count}\n" +
                                        $"Type:{c.Type}\n" +
                                        $"UTCtime:{c.UTCtime}\n" +
                                        $"isOutlier:{c.isOutlier}\n" +

                                        // Reference coordinates
                                        $"Nref:{c.Nref:F3} Eref:{c.Eref:F3} Href:{c.Href:F3}\n" +

                                        // Displacements
                                        $"dN:{c.dN:F3} dE:{c.dE:F3} dH:{c.dH:F3}\n" +

                                        // Corrections
                                        $"dNcor:{c.dNcor:F3} dEcor:{c.dEcor:F3} dHcor:{c.dHcor:F3}\n" +

                                        // Absolute coordinates
                                        $"N:{c.N:F3} E:{c.E:F3} H:{c.H:F3}\n" +

                                        // Mean coordinates
                                        $"Nmean:{c.MeanN:F3} Emean:{c.MeanE:F3} Hmean:{c.MeanH:F3}\n" +

                                        // Derived metrics
                                        $"dS:{c.dS:F3} dR:{c.dR:F3} dT:{c.dT:F3}\n" +

                                        // Time metrics
                                        $"Start UTC:{c.TimeBlockStartUTC} End UTC:{c.TimeBlockEndUTC}\n"
                                    );
                                }
                            }


                        // write coordinates to the workbook
                        bool writeSuccess = gnaSpreadsheetAPI.WriteCurrentCoordinatesToWorkbook(
                            env: env,
                            currentCoordinates: currentCoordinates);

                        if (!writeSuccess)
                        {
                            Console.WriteLine($"{strTab2}ERROR: Writing coordinates to workbook failed.");
                            Environment.Exit(1);
                        }


                        // write positional deltas to the workbook
                        Console.WriteLine($"{strTab2}Write deltas");
                        bool writeDeltasSuccess = gnaSpreadsheetAPI.WriteCurrentDeltasToWorkbook(
                            env: env,
                            currentCoordinates: currentCoordinates);

                        if (!writeDeltasSuccess)
                        {
                            Console.WriteLine($"{strTab2}ERROR: Writing deltas to workbook failed.");
                            Environment.Exit(1);
                        }



                        bool latestCoordsOk = gnaSpreadsheetAPI.prepareLatestCoordinatesWorksheet(env);

                        if (!latestCoordsOk)
                        {
                            Console.WriteLine($"{strTab2}ERROR: Preparing Latest Coordinates worksheet failed.");
                            Environment.Exit(1);
                        }

                        // Compute dR,dT and write to workbook if enabled

                        Console.WriteLine($"{strTab1}Compute Polar Vectors");
                        if (computedRdT)
                        {
                            Console.WriteLine($"{strTab2}Compute dRdTdH");
                            string[] parts = strReferenceLineTerminals.Split(',');

                            if (parts.Length != 4)
                            {
                                throw new ConfigurationErrorsException(
                                    $"{strTab2}ReferenceLineTerminalsEaNaEbNb must contain exactly four comma-separated values.");
                            }

                            double ea = double.Parse(parts[0], CultureInfo.InvariantCulture);
                            double na = double.Parse(parts[1], CultureInfo.InvariantCulture);
                            double eb = double.Parse(parts[2], CultureInfo.InvariantCulture);
                            double nb = double.Parse(parts[3], CultureInfo.InvariantCulture);

                            ReferenceLine2D referenceLine = new ReferenceLine2D(
                                a: new Coordinate2D(e: ea, n: na),
                                b: new Coordinate2D(e: eb, n: nb));

                            // write the terminal coordinates to historic worksheets
                            bool okTerminals = gnaSpreadsheetAPI.writeTerminalCoordinates(env, referenceLine);
                            if (!okTerminals)
                            {
                                Console.WriteLine($"{strTab2}ERROR: Writing reference line terminals failed.");
                                Environment.Exit(1);
                            }




                            if (debug)
                            {
                                Console.WriteLine("Reference Line Terminals (E,N):");

                                Console.WriteLine(
                                    $"    A  E:{referenceLine.A.E.ToString("F3", CultureInfo.InvariantCulture)}  " +
                                    $"N:{referenceLine.A.N.ToString("F3", CultureInfo.InvariantCulture)}");

                                Console.WriteLine(
                                    $"    B  E:{referenceLine.B.E.ToString("F3", CultureInfo.InvariantCulture)}  " +
                                    $"N:{referenceLine.B.N.ToString("F3", CultureInfo.InvariantCulture)}");
                            }

                            // compute dR and dT    
                            currentCoordinates = gnaSurvey.computedRdTdH(
                                referenceLine: referenceLine,
                                currentCoordinates: currentCoordinates);
                            if (debug)
                            {
                                Console.WriteLine();
                                Console.WriteLine("\nComputed Displacements (RT + Vertical):");
                                Console.WriteLine("--------------------------------------------------------------------------");

                                foreach (Points p in currentCoordinates)
                                {
                                    if (string.Equals(p.Count, "Missing", StringComparison.OrdinalIgnoreCase))
                                        continue;

                                    Console.WriteLine(
                                        $"Point: {p.Name,-15}  " +
                                        $"dR:{(p.dR.HasValue ? p.dR.Value.ToString("F5", CultureInfo.InvariantCulture) : "<null>"),8}  " +
                                        $"dT:{(p.dT.HasValue ? p.dT.Value.ToString("F5", CultureInfo.InvariantCulture) : "<null>"),8}  " +
                                        $"dH:{(p.dH.HasValue ? p.dH.Value.ToString("F5", CultureInfo.InvariantCulture) : "<null>"),8}  " +
                                        $"dHtot:{(p.dHtotal.HasValue ? p.dHtotal.Value.ToString("F5", CultureInfo.InvariantCulture) : "<null>"),8}");
                                }

                                Console.WriteLine("--------------------------------------------------------------------------");
                                Console.WriteLine();
                            }

                            // write dR, dT, dHtotal to the historic worksheets
                            Console.WriteLine($"{strTab2}Write dRdTdH");
                            bool writeRtSuccess = gnaSpreadsheetAPI.WriteCurrentRdTdHToWorkbook(
                                env: env,
                                currentCoordinates: currentCoordinates);

                            if (!writeRtSuccess)
                            {
                                Console.WriteLine($"{strTab2}ERROR: Writing dR/dT/dHtotal to workbook failed.");
                                Environment.Exit(1);
                            }

                            bool latestPolarOk = gnaSpreadsheetAPI.prepareLatestPolarWorksheet(env);

                            if (!latestPolarOk)
                            {
                                Console.WriteLine($"{strTab2}ERROR: Preparing Latest Polar Displacements worksheet failed.");
                                Environment.Exit(1);
                            }

                        }
                        else
                        {
                            Console.WriteLine($"{strTab2}dRdTdH not required");
                        }

                        Console.WriteLine($"{strTab1}Block Processing Done");
                    }

                    if (manualSurvey)
                    {
                        goto ThatsAllFolks;
                    }


                }
#endregion














ThatsAllFolks:
                FinishAndExit();
            }
            catch (Exception ex)
            {
                File.WriteAllText("fatal_crash.log", ex.ToString());
            }
        }
    
    
    
    }
}
