using BitMiracle.Docotic.Pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PDFtoExcel.Templates
{
    public class OXY
    {
        Dictionary<string, List<string>> _dictionaryOfColumns = new Dictionary<string, List<string>>();
        List<Dictionary<string, string>> _listOfColumns = new List<Dictionary<string, string>>();
        public OXY(Dictionary<string, List<string>> DictionaryOfColumns, List<Dictionary<string, string>> ListOfColumns)
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
            tableValues.Add("Lease & Contract Information Meter #", "");
            tableValues.Add("Lease & Contract Information Meter suf", "0");
            tableValues.Add("Lease & Contract Information State", "");
			_listOfColumns.Add(tableValues);

			tableValues = new Dictionary<string, string>();
			tableValues.Add("Lease & Contract Information Lease Name", "");
			_listOfColumns.Add(tableValues);

			tableValues = new Dictionary<string, string>();
			tableValues.Add("Lease & Contract Information Operator Name", "");
			_listOfColumns.Add(tableValues);

			tableValues = new Dictionary<string, string>();

			tableValues.Add("Lease & Contract Information Operator #", "");
            tableValues.Add("Lease & Contract Information Alloc Decimal", "");
            tableValues.Add("Lease & Contract Information Contract #", "");
			_listOfColumns.Add(tableValues);

			tableValues = new Dictionary<string, string>();
			tableValues.Add("Lease & Contract Information Ctr Pty Name", "");
			_listOfColumns.Add(tableValues);

			tableValues = new Dictionary<string, string>();
			tableValues.Add("Lease & Contract Information CTR Pty #", "");
			tableValues.Add("Lease & Contract Information Pressure Base", "");
			tableValues.Add("Lease & Contract Information BTU Basis", "");
			_listOfColumns.Add(tableValues);


            tableValues = new Dictionary<string, string>();
            tableValues.Add("Settlement Summary Residue Value", "");
            tableValues.Add("Settlement Summary Product Value", "");
            tableValues.Add("Settlement Summary Gross Value", "");
            tableValues.Add("Settlement Summary Services Fee", "");
            tableValues.Add("Settlement Summary Net Value", "");
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            foreach (var val in _dictionaryOfColumns["LiquidSettlementLabels"])
            {
                tableValues.Add("LiquidSettlement " + val + " Theoretical Volume", "");
                tableValues.Add("LiquidSettlement " + val + " Allocated Volume", "");
                tableValues.Add("LiquidSettlement " + val + " Shrink MMBTU", "");
                tableValues.Add("LiquidSettlement " + val + " POP%", "");
                tableValues.Add("LiquidSettlement " + val + " Settlement Volume", "");
                tableValues.Add("LiquidSettlement " + val + " TIK Volume", "");
                tableValues.Add("LiquidSettlement " + val + " Price", "");
                tableValues.Add("LiquidSettlement " + val + " Liquid Value", "");
            }
            _listOfColumns.Add(tableValues);

            tableValues = new Dictionary<string, string>();
            tableValues.Add("LiquidSettlement " + "Total Theoretical Volume", "");
            tableValues.Add("LiquidSettlement " + "Total Allocated Volume", "");
            tableValues.Add("LiquidSettlement " + "Total Shrink MMBTU", "");
            tableValues.Add("LiquidSettlement " + "Total Settlement Volume", "");
            tableValues.Add("LiquidSettlement " + "Total TIK Volume", "");
            tableValues.Add("LiquidSettlement " + "Total Liquid Value", "");
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
			tableValues.Add("Fees and Adjustments Total Value", "");
			_listOfColumns.Add(tableValues);


			tableValues = new Dictionary<string, string>();
			tableValues.Add("Analysis Methane Mol %", "");
			tableValues.Add("Analysis Methane GPM", null);
			tableValues.Add("Analysis Nitrogen Mol %", "");
			tableValues.Add("Analysis Nitrogen GPM", null);
			tableValues.Add("Analysis H2S Mol %", "");
			tableValues.Add("Analysis H2S GPM", null);
			tableValues.Add("Analysis Other Interts GPM", null);
			tableValues.Add("Analysis Other Interts Mol %", "");
			tableValues.Add("Analysis Ethane Mol %", "");
			tableValues.Add("Analysis Ethane GPM", "");
			tableValues.Add("Analysis Propane Mol %", "");
			tableValues.Add("Analysis Propane GPM", "");
			tableValues.Add("Analysis Iso Butane Mol %", "");
			tableValues.Add("Analysis Iso Butane GPM", "");
			tableValues.Add("Analysis Nor Butane Mol %", "");
			tableValues.Add("Analysis Nor Butane GPM", "");
			tableValues.Add("Analysis Natural Gas Mol %", "");
			tableValues.Add("Analysis Natural Gas GPM", "");
			tableValues.Add("Analysis Carbon Dioxide Mol %", "");
			tableValues.Add("Analysis Carbon Dioxide GPM", null);

			_listOfColumns.Add(tableValues);
			tableValues = new Dictionary<string, string>();
			tableValues.Add("Analysis Totals Mol%", "");
			tableValues.Add("Analysis Totals GPM", "");
			_listOfColumns.Add(tableValues);

			tableValues = new Dictionary<string, string>();
			tableValues.Add("Analysis WH BTU Factor", "");
			_listOfColumns.Add(tableValues);

			tableValues = new Dictionary<string, string>();
			tableValues.Add("ResidueSettlements MCF Net Delivery Point", "");
			tableValues.Add("ResidueSettlements MCF Unprocessed Plant Flare", "");
			tableValues.Add("ResidueSettlements MCF Shrink", "");
			tableValues.Add("ResidueSettlements MCF CO2", "");
			tableValues.Add("ResidueSettlements MCF Theoretical Residue", "");
			tableValues.Add("ResidueSettlements MCF Allocated Residue", "");
			tableValues.Add("ResidueSettlements MCF Purchased Fuels", "");
			tableValues.Add("ResidueSettlements MCF Residue Fuels", "");
			tableValues.Add("ResidueSettlements MCF Residue Flare", "");
			tableValues.Add("ResidueSettlements MCF Lease Return", "");
			tableValues.Add("ResidueSettlements MCF Residue TIK", "");
			tableValues.Add("ResidueSettlements MCF Settled Residue", "");
			_listOfColumns.Add(tableValues);

			tableValues = new Dictionary<string, string>();
			tableValues.Add("ResidueSettlements MMBTU Net Delivery Point", "");
			tableValues.Add("ResidueSettlements MMBTU Unprocessed Plant Flare", "");
			tableValues.Add("ResidueSettlements MMBTU Shrink", "");
			tableValues.Add("ResidueSettlements MMBTU CO2", "");
			tableValues.Add("ResidueSettlements MMBTU Theoretical Residue", "");
			tableValues.Add("ResidueSettlements MMBTU Allocated Residue", "");
			tableValues.Add("ResidueSettlements MMBTU Purchased Fuels", "");
			tableValues.Add("ResidueSettlements MMBTU Residue Fuels", "");
			tableValues.Add("ResidueSettlements MMBTU Residue Flare", "");
			tableValues.Add("ResidueSettlements MMBTU Lease Return", "");
			tableValues.Add("ResidueSettlements MMBTU Residue TIK", "");
			tableValues.Add("ResidueSettlements MMBTU Settled Residue", "");
			tableValues.Add("ResidueSettlements MMBTU Price", "");
			tableValues.Add("ResidueSettlements MMBTU Settlement %", "");
			tableValues.Add("ResidueSettlements MMBTU Resident Settlement Value", "");
			_listOfColumns.Add(tableValues);

			tableValues = new Dictionary<string, string>();
			tableValues.Add("VolumeInformation Gross Wellhead MCF", "");
			tableValues.Add("VolumeInformation Gross Wellhead MMBTU", "");
			tableValues.Add("VolumeInformation Gas Lift MCF", "");
			tableValues.Add("VolumeInformation Gas Lift MMBTU", "");
			tableValues.Add("VolumeInformation Wellhead Delivered MCF", "");
			tableValues.Add("VolumeInformation Wellhead Delivered MMBTU", "");
			tableValues.Add("VolumeInformation Inlet Adjustment MMBTU", "");
			tableValues.Add("VolumeInformation Inlet Adjustment MCF", "");
			tableValues.Add("VolumeInformation Net Delivered MCF", "");
			tableValues.Add("VolumeInformation Net Delivered MMBTU", "");
			_listOfColumns.Add(tableValues);

			tableValues = new Dictionary<string, string>();
			foreach (var val in _dictionaryOfColumns["PlantProductVolumesLabels"])
			{
				tableValues.Add("PlantProductVolumesLabels " + val + " Theoretical Volume", "");
				tableValues.Add("PlantProductVolumesLabels " + val + " Allocated Volume", "");
				tableValues.Add("PlantProductVolumesLabels " + val + " Shrink MMBTU", "");
			}
			_listOfColumns.Add(tableValues);

			tableValues = new Dictionary<string, string>();
			tableValues.Add("PlantProductVolumesLabels Totals Theoretical Volume", "");
			tableValues.Add("PlantProductVolumesLabels Totals Allocated Volume", "");
			tableValues.Add("PlantProductVolumesLabels Totals Shrink MMBTU", "");
			_listOfColumns.Add(tableValues);

			tableValues = new Dictionary<string, string>();
			tableValues.Add("PlantResidueSettlements MCF Gross Delivery Point", "");
			tableValues.Add("PlantResidueSettlements MCF Net Delivery Point", "");
			tableValues.Add("PlantResidueSettlements MCF Unprocessed Plant Flare", "");
			tableValues.Add("PlantResidueSettlements MCF Shrink", "");
			tableValues.Add("PlantResidueSettlements MCF CO2", "");
			tableValues.Add("PlantResidueSettlements MCF Theoretical Residue", "");
			tableValues.Add("PlantResidueSettlements MCF Allocated Residue", "");
			tableValues.Add("PlantResidueSettlements MCF Residue Fuels", "");
			tableValues.Add("PlantResidueSettlements MCF Residue Flare", "");
			_listOfColumns.Add(tableValues);

			tableValues = new Dictionary<string, string>();
			tableValues.Add("PlantResidueSettlements MMBTU Gross Delivery Point", "");
			tableValues.Add("PlantResidueSettlements MMBTU Net Delivery Point", "");
			tableValues.Add("PlantResidueSettlements MMBTU Unprocessed Plant Flare", "");
			tableValues.Add("PlantResidueSettlements MMBTU Shrink", "");
			tableValues.Add("PlantResidueSettlements MMBTU CO2", "");
			tableValues.Add("PlantResidueSettlements MMBTU Theoretical Residue", "");
			tableValues.Add("PlantResidueSettlements MMBTU Allocated Residue", "");
			tableValues.Add("PlantResidueSettlements MMBTU Residue Fuels", "");
			tableValues.Add("PlantResidueSettlements MMBTU Residue Flare", "");
			tableValues.Add("PlantResidueSettlements MMBTU Lease Return", "");
			tableValues.Add("PlantResidueSettlements MMBTU Settled Residue", "");
			_listOfColumns.Add(tableValues);
		}

		public void RunTemplate1(PdfPage pdf)
        {
            Dictionary<string, PdfRectangle> tableTemplates = new Dictionary<string, PdfRectangle>();
            tableTemplates.Add("LeaseAndContractInformationMeterAndState", new PdfRectangle(180, 87, 110, 20));
			tableTemplates.Add("LeaseAndContractInformationLeaseName", new PdfRectangle(300, 87, 60, 40));
			tableTemplates.Add("LeaseAndContractInformationOperatorName", new PdfRectangle(360, 87, 70, 40));
			tableTemplates.Add("LeaseAndContractInformationOperator#AndAllocDecimalAndCotnract#", new PdfRectangle(430, 87, 120, 30));
			tableTemplates.Add("LeaseAndContractInformationCtrPtyName", new PdfRectangle(565, 87, 70, 40));
			tableTemplates.Add("LeaseAndContractInformationCtrParty#AndPressureBaseAndBTUBasis", new PdfRectangle(630, 87, 130, 30));

			tableTemplates.Add("SettlementSummary", new PdfRectangle(100, 180, 50, 70));
            tableTemplates.Add("LiquidSettlementLabels", new PdfRectangle(180, 160, 80, 55));
			tableTemplates.Add("FeesAndAdjustmentsLabels", new PdfRectangle(233, 387, 90, 70));
			tableTemplates.Add("PlantProductVolumesLabels", new PdfRectangle(520, 390, 50, 65));
			ReadFromPDF1(pdf, tableTemplates["LiquidSettlementLabels"], "LiquidSettlementLabels"); // grabs the dynamic columns
			ReadFromPDF1(pdf, tableTemplates["FeesAndAdjustmentsLabels"], "FeesAndAdjustmentsLabels"); // grabs the dynamic columns
			ReadFromPDF1(pdf, tableTemplates["PlantProductVolumesLabels"], "PlantProductVolumesLabels"); // grabs the dynamic columns


			Initialize_listOfColumns();

            tableTemplates.Add("LiquidSettlementValues", new PdfRectangle(260, 160, 500, 60));
            tableTemplates.Add("LiquidSettlementTotals", new PdfRectangle(265, 240, 490, 220));
			tableTemplates.Add("FeesAndAdjustmentsValues", new PdfRectangle(330, 390, 160, 65));
			tableTemplates.Add("FeesAndAdjustmentsTotalValues", new PdfRectangle(445, 465, 50, 10));
			tableTemplates.Add("AnalysisValues", new PdfRectangle(100, 390, 100, 118));
			tableTemplates.Add("AnalysisTotalsValues", new PdfRectangle(95, 515, 105, 10));
			tableTemplates.Add("AnalysisWHBTUFactorValues", new PdfRectangle(100, 535, 40, 6));
			tableTemplates.Add("ResidueSettlementMCFValues", new PdfRectangle(50, 325, 560, 40));
			tableTemplates.Add("ResidueSettlementMMBTUValues", new PdfRectangle(50, 335, 695, 10));
			tableTemplates.Add("VolumeInformationValues", new PdfRectangle(75, 95, 80, 65));
			tableTemplates.Add("PlantProductVolumeValues", new PdfRectangle(575, 390, 185, 60));
			tableTemplates.Add("PlantProductVolumeTotalValues", new PdfRectangle(575, 470, 175, 5));
			tableTemplates.Add("PlantResidueSettlementMCFValues", new PdfRectangle(260, 540, 415, 10));
			tableTemplates.Add("PlantResidueSettlementMMBTUValues", new PdfRectangle(260, 550, 510, 10));


			//process each table and maps them back to the dictionary

			_listOfColumns[0] = ReadFromPDF(pdf, tableTemplates["LeaseAndContractInformationMeterAndState"], _listOfColumns[0]);
            _listOfColumns[1] = ReadFromPDF(pdf, tableTemplates["LeaseAndContractInformationLeaseName"], _listOfColumns[1],false, true,true);
			_listOfColumns[2] = ReadFromPDF(pdf, tableTemplates["LeaseAndContractInformationOperatorName"], _listOfColumns[2], false, true, true);
			_listOfColumns[3] = ReadFromPDF(pdf, tableTemplates["LeaseAndContractInformationOperator#AndAllocDecimalAndCotnract#"], _listOfColumns[3]);
			_listOfColumns[4] = ReadFromPDF(pdf, tableTemplates["LeaseAndContractInformationCtrPtyName"], _listOfColumns[4], false, true, true);
			_listOfColumns[5] = ReadFromPDF(pdf, tableTemplates["LeaseAndContractInformationCtrParty#AndPressureBaseAndBTUBasis"], _listOfColumns[5]);
			
			_listOfColumns[6] = ReadFromPDF(pdf, tableTemplates["SettlementSummary"], _listOfColumns[6], true, true);
			_listOfColumns[7] = ReadFromPDF(pdf, tableTemplates["LiquidSettlementValues"], _listOfColumns[7], false, true);
			_listOfColumns[8] = ReadFromPDF(pdf, tableTemplates["LiquidSettlementTotals"], _listOfColumns[8], false, true);
			_listOfColumns[9] = ReadFromPDF(pdf, tableTemplates["FeesAndAdjustmentsValues"], _listOfColumns[9], false, true);
			_listOfColumns[10] = ReadFromPDF(pdf, tableTemplates["FeesAndAdjustmentsTotalValues"], _listOfColumns[10], false, true);
			_listOfColumns[11] = ReadFromPDF(pdf, tableTemplates["AnalysisValues"], _listOfColumns[11], false, true);
			_listOfColumns[12] = ReadFromPDF(pdf, tableTemplates["AnalysisTotalsValues"], _listOfColumns[12], false, true);
			_listOfColumns[13] = ReadFromPDF(pdf, tableTemplates["AnalysisWHBTUFactorValues"], _listOfColumns[13], false, true);
			_listOfColumns[14] = ReadFromPDF(pdf, tableTemplates["ResidueSettlementMCFValues"], _listOfColumns[14], false, true);
			_listOfColumns[15] = ReadFromPDF(pdf, tableTemplates["ResidueSettlementMMBTUValues"], _listOfColumns[15], false, true);
			_listOfColumns[16] = ReadFromPDF(pdf, tableTemplates["VolumeInformationValues"], _listOfColumns[16], true, true);
			_listOfColumns[17] = ReadFromPDF(pdf, tableTemplates["PlantProductVolumeValues"], _listOfColumns[17], false, true);
			_listOfColumns[18] = ReadFromPDF(pdf, tableTemplates["PlantProductVolumeTotalValues"], _listOfColumns[18], false, true);
			_listOfColumns[19] = ReadFromPDF(pdf, tableTemplates["PlantResidueSettlementMCFValues"], _listOfColumns[19], false, true);
			_listOfColumns[20] = ReadFromPDF(pdf, tableTemplates["PlantResidueSettlementMMBTUValues"], _listOfColumns[20], false, true);
		}

	}
}
