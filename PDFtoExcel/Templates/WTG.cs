using BitMiracle.Docotic.Pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PDFtoExcel.Templates
{
    public class WTG
    {
        Dictionary<string, List<string>> _dictionaryOfColumns = new Dictionary<string, List<string>>();
        List<Dictionary<string, string>> _listOfColumns = new List<Dictionary<string, string>>();
        public WTG(Dictionary<string, List<string>> DictionaryOfColumns, List<Dictionary<string, string>> ListOfColumns)
        {
            _dictionaryOfColumns = DictionaryOfColumns;
            _listOfColumns = ListOfColumns;
        }

        public List<Dictionary<string, string>> ListOfColumns { get { return _listOfColumns; } }
        private void ReadFromPDF1(PdfPage pdf, PdfRectangle rectangle, string columnName)
        {
            var options = new PdfTextExtractionOptions
            {
                Rectangle = rectangle,
                WithFormatting = false
            };
            string areaText = pdf.GetText(options);
            _dictionaryOfColumns.Add(columnName, areaText.Split("\r\n").ToList());

        }

        private static Dictionary<string, string> ReadFromPDF(PdfPage pdf, PdfRectangle rectangle, Dictionary<string, string> columns, bool withFormatting = true, bool isNumberOnly = false, bool newLineSplit = false)
        {
            var options = new PdfTextExtractionOptions
            {
                Rectangle = rectangle,
                WithFormatting = withFormatting // grabs the data with proper spacing and formatting, if false it just grabs the data without it
            };
            string areaText = pdf.GetText(options);
            string[] splitText = null;
            if (!withFormatting)
                areaText = areaText.Replace("\r\n", " "); // replaces the newline into space so we can use it as a delimiter
            if (withFormatting && isNumberOnly)
            { // special logic to force the parser to go left-to-right instead of top-to-bottom in processing parsed data
                areaText = areaText.Trim();
                Regex regex = new Regex("[ ]{2,}", RegexOptions.None);
                areaText = areaText.Replace("\r\n", " ");
                areaText = regex.Replace(areaText, " ");
            }
            if (newLineSplit) // some tables need this as a delimiter
                splitText = areaText.Split("\r\n");
            else
                splitText = areaText.Split(" "); // use whitespace as delimiter
            List<string> keys = new List<string>(columns.Keys);


            int ctr = 0;
            bool wasEmpty = false;

            /*
             * logic in processing parsed data
             * very convoluted, could probably be improved
             */
            foreach (var key in keys)
            {

                bool wordFound = false;
                int whitespaceCtr = 0;
                string word = "";
                for (int i = ctr; i < splitText.Length; i++)
                {
                    if (!String.IsNullOrEmpty(columns[key]))
                    {
                        if (String.IsNullOrEmpty(splitText[i]))
                        {
                            wasEmpty = true;
                        }
                        break;
                    }

                    if (wasEmpty)
                    {
                        i++;
                        ctr++;
                        wasEmpty = false;
                    }
                    if (!isNumberOnly)
                    {
                        if (!String.IsNullOrEmpty(splitText[i]))
                        {
                            if (string.IsNullOrEmpty(word))
                                word = splitText[i];
                            else
                            {
                                word = word + " " + splitText[i];
                            }
                            wordFound = true;
                            whitespaceCtr = 0;
                            if (newLineSplit)
                            {
                                ctr++;
                                break;
                            }
                        }
                        else
                        {
                            if (wordFound)
                                if (whitespaceCtr < 1)
                                    whitespaceCtr++;
                                else
                                    break;
                        }
                    }
                    else
                    {
                        if (columns[key] == null)
                            word = "0";
                        else
                        {
                            word = splitText[i];
                            ctr++;
                        }
                        break;
                    }
                    ctr++;
                }
                columns[key] = word;
            }

            return columns;
        }

        /*
         Grabs dynamic and sets static columns so we can fill them with the parsed data later on
         */
        private void Initialize_listOfColumns()
        {
            _listOfColumns = new List<Dictionary<string, string>>();
            Dictionary<string, string> tableValues = new Dictionary<string, string>();
            tableValues.Add("Lease & Contract Information Facility Name", "");
            tableValues.Add("Lease & Contract Information Production Date", "");
            tableValues.Add("Lease & Contract Information Accounting Date", "");
            tableValues.Add("Lease & Contract Information Lease Name", "");
            tableValues.Add("Lease & Contract Information Allocation Decimal", "");
            tableValues.Add("Lease & Contract Information Meter Number", "");
            tableValues.Add("Lease & Contract Information State", "");
            tableValues.Add("Lease & Contract Information County", "");
            tableValues.Add("Lease & Contract Information Contract Number", "");
            tableValues.Add("Lease & Contract Information Pressure Base", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("Settlement Summary Residue Value", "");
            tableValues.Add("Settlement Summary Liquid Value", "");
            tableValues.Add("Settlement Summary Gross Value", "");
            tableValues.Add("Settlement Summary Fees & Adjustments", "");
            tableValues.Add("Settlement Summary Tax", "");
            tableValues.Add("Settlement Summary Tax Reimbursment", "");
            tableValues.Add("Settlement Summary Net Value", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            foreach (var val in _dictionaryOfColumns["LiquidSettlementLabels"])
            {
                tableValues.Add("LiquidSettlement " + val + " Allocated Gallons", "");
                tableValues.Add("LiquidSettlement " + val + " Settled MMBTU", "");
                tableValues.Add("LiquidSettlement " + val + " Contract %", "");
                tableValues.Add("LiquidSettlement " + val + " Settlement Gallons", "");
                tableValues.Add("LiquidSettlement " + val + " Price", "");
                tableValues.Add("LiquidSettlement " + val + " Liquid Value", "");
            }
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("LiquidSettlement " + "Total Allocated Gallons", "");
            tableValues.Add("LiquidSettlement " + "Total Settled MMBTU", "");
            tableValues.Add("LiquidSettlement " + "Total Settlement Gallons", "");
            tableValues.Add("LiquidSettlement " + "Total Liquid Value", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            foreach (var val in _dictionaryOfColumns["WellHeadInformationLabels"])
            {
                if (!String.IsNullOrEmpty(val))
                {
                    if (val == "Fuel (Off  System)")
                    {
                        tableValues.Add("Wellhead Information " + val + " Mcf", "0");
                        tableValues.Add("Wellhead Information " + val + " MMBTU", "0");
                    }
                    else if (val == "Wellhead Btu Factor:")
                    {
                        tableValues.Add("Wellhead Information " + val + " Mcf", "0");
                        tableValues.Add("Wellhead Information " + val + " MMBTU", "");
                    }
                    else
                    {
                        tableValues.Add("Wellhead Information " + val + " Mcf", "");
                        tableValues.Add("Wellhead Information " + val + " MMBTU", "");

                    }
                }
            }
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("Residue Allocation Net Delivered Mcf", "");
            tableValues.Add("Residue Allocation Net Delivered MBBTU", "");
            tableValues.Add("Residue Allocation Shrink Mcf", "");
            tableValues.Add("Residue Allocation Shrink MBBTU", "");
            tableValues.Add("Residue Allocation Plant Fuel Mcf", "");
            tableValues.Add("Residue Allocation Plant Fuel MBBTU", "");
            tableValues.Add("Residue Allocation Actual Residue Mcf", "");
            tableValues.Add("Residue Allocation Actual Residue MBBTU", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("Residue Settlement Contract %", "");
            tableValues.Add("Residue Settlement Settlement Residue", "");
            tableValues.Add("Residue Settlement Price", "");
            tableValues.Add("Residue Settlement Residue Value", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            foreach (var val in _dictionaryOfColumns["FeesAndAdjustmentsLabels"])
            {
                tableValues.Add("Fees and Adjustments " + val + " Basis", "");
                tableValues.Add("Fees and Adjustments " + val + " Rate", "");
                tableValues.Add("Fees and Adjustments " + val + " Value", "");
            }
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("Analysis Nitrogen Mol %", "");
            tableValues.Add("Analysis Nitrogen GPM", null);
            tableValues.Add("Analysis Carbon Dioxide Mol %", "");
            tableValues.Add("Analysis Carbon Dioxide GPM", null);
            tableValues.Add("Analysis H2S Mol %", "");
            tableValues.Add("Analysis H2S GPM", null);
            tableValues.Add("Analysis Other Interts Mol %", "");
            tableValues.Add("Analysis Other Interts GPM", null);
            tableValues.Add("Analysis Methane Mol %", "");
            tableValues.Add("Analysis Methane GPM", null);
            tableValues.Add("Analysis Ethane Mol %", "");
            tableValues.Add("Analysis Ethane GPM", "");
            tableValues.Add("Analysis Propane Mol %", "");
            tableValues.Add("Analysis Propane GPM", "");
            tableValues.Add("Analysis Iso Butane Mol %", "");
            tableValues.Add("Analysis Iso Butane GPM", "");
            tableValues.Add("Analysis Nor Butane Mol %", "");
            tableValues.Add("Analysis Nor Butane GPM", "");
            tableValues.Add("Analysis Iso Pentane Mol %", "");
            tableValues.Add("Analysis Iso Pentane GPM", "");
            tableValues.Add("Analysis Nor Pentane Mol %", "");
            tableValues.Add("Analysis Nor Pentane GPM", "");
            tableValues.Add("Analysis Hexane Mol %", "");
            tableValues.Add("Analysis Hexane GPM", "");
            tableValues.Add("Analysis Total Mol %", "");
            tableValues.Add("Analysis Total GPM", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("Analysis H2S PPM", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("Analysis Specific Gravity", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("Plant Contacts Accounting Name", "");
            tableValues.Add("Plant Contacts Accounting Number", "");
            tableValues.Add("Plant Contacts Contracts Name", "");
            tableValues.Add("Plant Contacts Contracts Number", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("Operator Nm", "");
            tableValues.Add("Operator ID", "");
            tableValues.Add("Ctr Pty Nm", "");
            tableValues.Add("Ctr Pty ID", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("Comments", "");
            _listOfColumns.Add(tableValues);

        }

        /*
         Grabs dynamic and sets static columns so we can fill them with the parsed data later on
         */
        private void Initialize_listOfColumns2()
        {
            _listOfColumns = new List<Dictionary<string, string>>();
            Dictionary<string, string> tableValues = new Dictionary<string, string>();
            tableValues.Add("Lease & Contract Information Facility Name", "");
            tableValues.Add("Lease & Contract Information Production Date", "");
            tableValues.Add("Lease & Contract Information Accounting Date", "");
            tableValues.Add("Lease & Contract Information Lease Name", "");
            tableValues.Add("Lease & Contract Information Allocation Decimal", "");
            tableValues.Add("Lease & Contract Information Meter Number", "");
            tableValues.Add("Lease & Contract Information State", "");
            tableValues.Add("Lease & Contract Information County", "");
            tableValues.Add("Lease & Contract Information Contract Number", "");
            tableValues.Add("Lease & Contract Information Pressure Base", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("Settlement Summary Residue Value", "");
            tableValues.Add("Settlement Summary Liquid Value", "");
            tableValues.Add("Settlement Summary Settled Value", "");
            tableValues.Add("Settlement Summary Fees & Adjustments", "");
            tableValues.Add("Settlement Summary Net Value Before Taxes", "");
            tableValues.Add("Settlement Summary Tax", "");
            tableValues.Add("Settlement Summary Tax Reimbursment", "");
            tableValues.Add("Settlement Summary Net Value", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            foreach (var val in _dictionaryOfColumns["LiquidSettlementLabels"])
            {
                tableValues.Add("LiquidSettlement " + val + " Plant Recovery %", "");
                tableValues.Add("LiquidSettlement " + val + " GPM", "");
                tableValues.Add("LiquidSettlement " + val + " Recovered Gallons %", "");
                tableValues.Add("LiquidSettlement " + val + " Shrink", "");
                tableValues.Add("LiquidSettlement " + val + " Contract %", "");
                tableValues.Add("LiquidSettlement " + val + " Settlement Gallons", "");
                tableValues.Add("LiquidSettlement " + val + " Price", "");
                tableValues.Add("LiquidSettlement " + val + " Liquid Value", "");
            }
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("LiquidSettlement " + "Total GPM", "");
            tableValues.Add("LiquidSettlement " + "Total Recovered Gallons", "");
            tableValues.Add("LiquidSettlement " + "Total Shrink", "");
            tableValues.Add("LiquidSettlement " + "Total Settlement Gallons", "");
            tableValues.Add("LiquidSettlement " + "Total Liquid Value", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("Wellhead Information Wellhead: Mcf", "");
            tableValues.Add("Wellhead Information Split Decimal: Mcf", "");
            tableValues.Add("Wellhead Information Owner Share: Mcf", "");
            tableValues.Add("Wellhead Information Fuel (On System) Mcf", "");
            tableValues.Add("Wellhead Information Fuel (Off System) Mcf", "");

            tableValues.Add("Wellhead Information Wellhead: MMBTU", "");
            tableValues.Add("Wellhead Information Split Decimal: MMBTU", "");
            tableValues.Add("Wellhead Information Owner Share: MMBTU", "");
            tableValues.Add("Wellhead Information Fuel (On System) MMBTU", "");
            tableValues.Add("Wellhead Information Fuel (Off System) MMBTU", "");

            tableValues.Add("Wellhead Information Wellhead Btu Factor: Mcf", "");
            tableValues.Add("Wellhead Information Net Delivered: Mcf", "");
            tableValues.Add("Wellhead Information Net Delivered: MMBTU", "");
            tableValues.Add("Wellhead Information Wellhead Btu Factor: MMBTU", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("Residue Allocation Net Delivered Mcf", "");
            tableValues.Add("Residue Allocation Shrink Mcf", "");
            tableValues.Add("Residue Allocation Plant Fuel Mcf", "");
            tableValues.Add("Residue Allocation Actual Residue Mcf", "");
            tableValues.Add("Residue Allocation Net Delivered MBBTU", "");
            tableValues.Add("Residue Allocation Shrink MBBTU", "");
            tableValues.Add("Residue Allocation Plant Fuel MBBTU", "");
            tableValues.Add("Residue Allocation Actual Residue MBBTU", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("Residue Settlement Contract %", "");
            tableValues.Add("Residue Settlement Settlement Residue", "");
            tableValues.Add("Residue Settlement Price", "");
            tableValues.Add("Residue Settlement Residue Value", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            foreach (var val in _dictionaryOfColumns["FeesAndAdjustmentsLabels"])
            {
                tableValues.Add("Fees And Adjustments " + val + " Basis", "");
                tableValues.Add("Fees And Adjustments " + val + " Rate", "");
                tableValues.Add("Fees And Adjustments " + val + " Value", "");
            }

            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("Analysis Nitrogen Mol %", "");
            tableValues.Add("Analysis Nitrogen GPM", null);
            tableValues.Add("Analysis Carbon Dioxide Mol %", "");
            tableValues.Add("Analysis Carbon Dioxide GPM", null);
            tableValues.Add("Analysis H2S Mol %", "");
            tableValues.Add("Analysis H2S GPM", null);
            tableValues.Add("Analysis Other Interts Mol %", "");
            tableValues.Add("Analysis Other Interts GPM", null);
            tableValues.Add("Analysis Methane Mol %", "");
            tableValues.Add("Analysis Methane GPM", null);
            tableValues.Add("Analysis Ethane Mol %", "");
            tableValues.Add("Analysis Ethane GPM", "");
            tableValues.Add("Analysis Propane Mol %", "");
            tableValues.Add("Analysis Propane GPM", "");
            tableValues.Add("Analysis Iso Butane Mol %", "");
            tableValues.Add("Analysis Iso Butane GPM", "");
            tableValues.Add("Analysis Nor Butane Mol %", "");
            tableValues.Add("Analysis Nor Butane GPM", "");
            tableValues.Add("Analysis Iso Pentane Mol %", "");
            tableValues.Add("Analysis Iso Pentane GPM", "");
            tableValues.Add("Analysis Nor Pentane Mol %", "");
            tableValues.Add("Analysis Nor Pentane GPM", "");
            tableValues.Add("Analysis Hexane Mol %", "");
            tableValues.Add("Analysis Hexane GPM", "");
            tableValues.Add("Analysis Total Mol %", "");
            tableValues.Add("Analysis Total GPM", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("Analysis H2S PPM", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("Analysis Specific Gravity", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("Plant Contacts Accounting Name", "");
            tableValues.Add("Plant Contacts Accounting Number", "");
            tableValues.Add("Plant Contacts Contracts Name", "");
            tableValues.Add("Plant Contacts Contracts Number", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("Operator Nm", "");
            tableValues.Add("Operator ID", "");
            tableValues.Add("Ctr Pty Nm", "");
            tableValues.Add("Ctr Pty ID", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("Comments", "");
            _listOfColumns.Add(tableValues);
        }
        public void RunTemplate1(PdfPage pdf)
        {
            Dictionary<string, PdfRectangle> tableTemplates = new Dictionary<string, PdfRectangle>();
            tableTemplates.Add("LeaseAndContractInformation", new PdfRectangle(200, 75, 550, 10));
            tableTemplates.Add("SettlementSummary", new PdfRectangle(200, 130, 550, 10));
            tableTemplates.Add("LiquidSettlementLabels", new PdfRectangle(220, 195, 80, 60));
            tableTemplates.Add("WellHeadInformationLabels", new PdfRectangle(30, 180, 70, 100));
            tableTemplates.Add("FeesAndAdjustmentsLabels", new PdfRectangle(30, 415, 70, 150));
            ReadFromPDF1(pdf, tableTemplates["LiquidSettlementLabels"], "LiquidSettlementLabels"); // grabs the dynamic columns
            ReadFromPDF1(pdf, tableTemplates["WellHeadInformationLabels"], "WellHeadInformationLabels"); // grabs the dynamic columns
            ReadFromPDF1(pdf, tableTemplates["FeesAndAdjustmentsLabels"], "FeesAndAdjustmentsLabels"); // grabs the dynamic columns

            Initialize_listOfColumns();

            tableTemplates.Add("LiquidSettlementValues", new PdfRectangle(340, 195, 420, 90));
            tableTemplates.Add("LiquidSettlementTotals", new PdfRectangle(340, 280, 420, 30));
            tableTemplates.Add("WellHeadInformationValues", new PdfRectangle(110, 180, 80, 110));
            tableTemplates.Add("ResidueAllocationValues", new PdfRectangle(190, 340, 110, 40));
            tableTemplates.Add("ResidueSettlementValues", new PdfRectangle(315, 345, 210, 40));
            tableTemplates.Add("FeesAndAdjustmentsValues", new PdfRectangle(110, 420, 115, 140));
            tableTemplates.Add("AnalysisValues", new PdfRectangle(290, 420, 100, 120));
            tableTemplates.Add("AnalysisH2SPPMValues", new PdfRectangle(390, 425, 25, 100));
            tableTemplates.Add("AnalysisSpecificGravityValues", new PdfRectangle(390, 555, 25, 10));
            tableTemplates.Add("PlantContactsValues", new PdfRectangle(580, 410, 140, 60));
            tableTemplates.Add("OperatorAndCtrInfoValues", new PdfRectangle(70, 80, 110, 70));
            tableTemplates.Add("CommentsValues", new PdfRectangle(450, 495, 320, 490));


            //process each table and maps them back to the dictionary
            _listOfColumns[0] = ReadFromPDF(pdf, tableTemplates["LeaseAndContractInformation"], _listOfColumns[0]);
            _listOfColumns[1] = ReadFromPDF(pdf, tableTemplates["SettlementSummary"], _listOfColumns[1]);
            _listOfColumns[2] = ReadFromPDF(pdf, tableTemplates["LiquidSettlementValues"], _listOfColumns[2]);
            _listOfColumns[3] = ReadFromPDF(pdf, tableTemplates["LiquidSettlementTotals"], _listOfColumns[3]);
            _listOfColumns[4] = ReadFromPDF(pdf, tableTemplates["WellHeadInformationValues"], _listOfColumns[4], false, true);
            _listOfColumns[5] = ReadFromPDF(pdf, tableTemplates["ResidueAllocationValues"], _listOfColumns[5], false, true);
            _listOfColumns[6] = ReadFromPDF(pdf, tableTemplates["ResidueSettlementValues"], _listOfColumns[6], false, true);
            _listOfColumns[7] = ReadFromPDF(pdf, tableTemplates["FeesAndAdjustmentsValues"], _listOfColumns[7], false, true);
            _listOfColumns[8] = ReadFromPDF(pdf, tableTemplates["AnalysisValues"], _listOfColumns[8], false, true);
            _listOfColumns[9] = ReadFromPDF(pdf, tableTemplates["AnalysisH2SPPMValues"], _listOfColumns[9], false, true);
            _listOfColumns[10] = ReadFromPDF(pdf, tableTemplates["AnalysisSpecificGravityValues"], _listOfColumns[10], false, true);
            _listOfColumns[11] = ReadFromPDF(pdf, tableTemplates["PlantContactsValues"], _listOfColumns[11], true, false, true);
            _listOfColumns[12] = ReadFromPDF(pdf, tableTemplates["OperatorAndCtrInfoValues"], _listOfColumns[12], true, false, true);
            _listOfColumns[13] = ReadFromPDF(pdf, tableTemplates["CommentsValues"], _listOfColumns[13]);
        }

        public void RunTemplate2(PdfPage pdf)
        {
            Dictionary<string, PdfRectangle> tableTemplates = new Dictionary<string, PdfRectangle>();

            tableTemplates.Add("LeaseAndContractInformation", new PdfRectangle(200, 78, 580, 15));
            tableTemplates.Add("SettlementSummary", new PdfRectangle(200, 130, 550, 10));
            tableTemplates.Add("LiquidSettlementLabels", new PdfRectangle(200, 200, 80, 60));
            tableTemplates.Add("FeesAndAdjustmentsLabels", new PdfRectangle(30, 415, 70, 150));
            ReadFromPDF1(pdf, tableTemplates["LiquidSettlementLabels"], "LiquidSettlementLabels"); // grabs the dynamic columns
            ReadFromPDF1(pdf, tableTemplates["FeesAndAdjustmentsLabels"], "FeesAndAdjustmentsLabels"); // grabs the dynamic columns

            Initialize_listOfColumns2();

            tableTemplates.Add("LiquidSettlementValues", new PdfRectangle(290, 200, 470, 75));
            tableTemplates.Add("LiquidSettlementTotals", new PdfRectangle(340, 280, 420, 30));
            tableTemplates.Add("WellHeadInformationValues", new PdfRectangle(100, 185, 100, 120));
            tableTemplates.Add("ResidueAllocationValues", new PdfRectangle(190, 340, 110, 40));
            tableTemplates.Add("ResidueSettlementValues", new PdfRectangle(315, 345, 210, 40));
            tableTemplates.Add("FeesAndAdjustmentsValues", new PdfRectangle(120, 425, 115, 140));
            tableTemplates.Add("AnalysisValues", new PdfRectangle(310, 425, 100, 120));
            tableTemplates.Add("AnalysisH2SPPMValues", new PdfRectangle(360, 555, 20, 10));
            tableTemplates.Add("AnalysisSpecificGravityValues", new PdfRectangle(350, 565, 20, 10));
            tableTemplates.Add("PlantContactsValues", new PdfRectangle(580, 410, 140, 60));
            tableTemplates.Add("OperatorAndCtrInfoValues", new PdfRectangle(60, 95, 110, 70));
            tableTemplates.Add("CommentsValues", new PdfRectangle(390, 495, 320, 490));


            //process each table and maps them back to the dictionary
            _listOfColumns[0] = ReadFromPDF(pdf, tableTemplates["LeaseAndContractInformation"], _listOfColumns[0]);
            _listOfColumns[1] = ReadFromPDF(pdf, tableTemplates["SettlementSummary"], _listOfColumns[1]);
            _listOfColumns[2] = ReadFromPDF(pdf, tableTemplates["LiquidSettlementValues"], _listOfColumns[2], false, true);
            _listOfColumns[3] = ReadFromPDF(pdf, tableTemplates["LiquidSettlementTotals"], _listOfColumns[3]);
            _listOfColumns[4] = ReadFromPDF(pdf, tableTemplates["WellHeadInformationValues"], _listOfColumns[4], false, true);
            _listOfColumns[5] = ReadFromPDF(pdf, tableTemplates["ResidueAllocationValues"], _listOfColumns[5], false, true);
            _listOfColumns[6] = ReadFromPDF(pdf, tableTemplates["ResidueSettlementValues"], _listOfColumns[6], false, true);
            _listOfColumns[7] = ReadFromPDF(pdf, tableTemplates["FeesAndAdjustmentsValues"], _listOfColumns[7], true, true);
            _listOfColumns[8] = ReadFromPDF(pdf, tableTemplates["AnalysisValues"], _listOfColumns[8], false, true);
            _listOfColumns[9] = ReadFromPDF(pdf, tableTemplates["AnalysisH2SPPMValues"], _listOfColumns[9], false, true);
            _listOfColumns[10] = ReadFromPDF(pdf, tableTemplates["AnalysisSpecificGravityValues"], _listOfColumns[10], false, true);
            _listOfColumns[11] = ReadFromPDF(pdf, tableTemplates["PlantContactsValues"], _listOfColumns[11], true, false, true);
            _listOfColumns[12] = ReadFromPDF(pdf, tableTemplates["OperatorAndCtrInfoValues"], _listOfColumns[12], true, false, true);
            _listOfColumns[13] = ReadFromPDF(pdf, tableTemplates["CommentsValues"], _listOfColumns[13]);
        }
    }
}
