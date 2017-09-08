using Bytescout.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Data_generator_for_REFLET
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileStream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            
            openFileDialog1.Filter = "Excel |*.xlsx";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.Title = "Select the file with data";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                try
                {
                    if ((myStream = openFileDialog1.OpenFile() as FileStream) != null)
                    {
                        using (myStream)
                        {
                            var templateFilePath = myStream.Name;
                            var path = templateFilePath.Substring(0, templateFilePath.LastIndexOf("\\"));
                            var fileName = "Data for REFLET.xlsx";

                            string file_path = Path.Combine(path, fileName);

                            Spreadsheet document = new Spreadsheet();
                            document.LoadFromStream(myStream);

                            Worksheet new_sheet = document.Workbook.Worksheets.Add("Logs - data");
                            Worksheet treated_sheet = document.Workbook.Worksheets.Add("Treated data");
                            Worksheet template_sheet = document.Workbook.Worksheets.ByName("Results");
                            Worksheet matching_sheet = document.Workbook.Worksheets.ByName("Matching");
                            Worksheet data_sheet = document.Workbook.Worksheets.ByName("Data");
                            Worksheet offers_to_link_sheet = document.Workbook.Worksheets.ByName("Offers to link");
                            Worksheet matching_contract_site_sheet = document.Workbook.Worksheets.ByName("Matching Contract-Site");

                            //new_sheet.Cell(0, 0).Value = "From matching";
                            new_sheet.Cell(0, 0).Value = "Columns not fount in data sheet";
                            new_sheet.Cell(0, 1).Value = "Other";

                            //Counters
                            int c_remarks = 1;

                            //If offers_to_link_sheet exists, the template concerns the Sites
                            if (offers_to_link_sheet != null && matching_contract_site_sheet != null)
                            {
                                var contractCellResults = template_sheet.Find("Nom Global Report", false, false, false);
                                var siteNameCellResults = template_sheet.Find("Nom du site", false, false, false);

                                if (contractCellResults != null)
                                {
                                    var last_used_row_results = template_sheet.NotEmptyRowMax;

                                    if (last_used_row_results > 1)
                                    {
                                        if (matching_contract_site_sheet.Find("Nom du contrat", false, false, false) != null)
                                        {
                                            for (int i = 1; i < last_used_row_results + 1; i++)
                                            {
                                                var col_gr = contractCellResults.GetAddress().Column;
                                                var grName = template_sheet.Cell(i, col_gr).Value;

                                                if (grName != null && grName.ToString() != "")
                                                {
                                                    //Chercher dans matching contract-site
                                                    var siteCellContractSite = matching_contract_site_sheet.Find(grName.ToString(), false, false, false);
                                                    if (siteCellContractSite != null && siteCellContractSite.Value.ToString() != "")
                                                    {
                                                        var contractCellContractSiteValue = matching_contract_site_sheet.Cell(siteCellContractSite.GetAddress().Row, matching_contract_site_sheet.Find("Nom du contrat", false, false, false).GetAddress().Column).Value.ToString();

                                                        //If the site has a matched contract in Contract-Site sheet, continue to Offers to link sheet
                                                        if (contractCellContractSiteValue != "")
                                                        {
                                                            var contractColOffersToLink = offers_to_link_sheet.Find("Nom du contrat", false, false, false);

                                                            //Check if "Nom du contrat" column exists in Offers to link sheet
                                                            if (contractColOffersToLink != null)
                                                            {
                                                                var contractCellOffersToLink = offers_to_link_sheet.Find(contractCellContractSiteValue, false, false, false);

                                                                if (contractCellOffersToLink != null)
                                                                {
                                                                    int col_contract_offers_To_Link = contractColOffersToLink.GetAddress().Column;
                                                                    int row_contract_offers_To_Link = contractCellOffersToLink.GetAddress().Row;
                                                                    int nbColsOffersToLink = offers_to_link_sheet.UsedRangeColumnMax;

                                                                    int colNb = new_sheet.UsedRangeColumnMax + 1;

                                                                    for (int j = 0; j < nbColsOffersToLink + 1; j ++)
                                                                    {
                                                                        var colName = offers_to_link_sheet.Cell(0, j).Value;

                                                                        if (colName != null && colName.ToString() != "" && colName.ToString() != "Nom du contrat")
                                                                        {
                                                                            var cellInLoop = template_sheet.Find(colName.ToString(), false, false, false);

                                                                            if (cellInLoop != null)
                                                                            {
                                                                                int colCellInLoop = cellInLoop.GetAddress().Column;
                                                                                var cellValueOffertsToLink = offers_to_link_sheet.Cell(row_contract_offers_To_Link, j).Value;
                                                                                template_sheet.Cell(i, colCellInLoop).Value = cellValueOffertsToLink;
                                                                            }
                                                                            else
                                                                            {
                                                                                new_sheet.Cell(c_remarks, colNb).Value = "Colonne '" + colName.ToString() + "' non trouvé dans Results";
                                                                                colNb++;
                                                                            }
                                                                        }
                                                                        else
                                                                        {
                                                                            if (colName != null && colName.ToString() != "Nom du contrat")
                                                                            {
                                                                                new_sheet.Cell(c_remarks, colNb).Value = "Contrat '" + contractCellContractSiteValue + "' non trouvé dans 'Offers to link'";
                                                                                colNb++;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    new_sheet.Cell(c_remarks, 2).Value = "Contrat '" + contractCellContractSiteValue + "' non trouvé dans 'Offers to link'";
                                                                    c_remarks++;
                                                                }
                                                            }
                                                            else
                                                            {
                                                                new_sheet.Cell(c_remarks, 2).Value = "Colonne 'Nom du contrat' non trouvé dans 'Offers to link'";
                                                                c_remarks++;
                                                            }
                                                        }
                                                        else
                                                        {
                                                            new_sheet.Cell(c_remarks, 2).Value = "Aucune correspondance pour le site : " + siteCellContractSite.Value.ToString();
                                                            c_remarks++;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        new_sheet.Cell(c_remarks, 2).Value = "'" + grName.ToString() + "' non trouvé dans Matching contract-site ou champ vide";
                                                        c_remarks++;
                                                    }
                                                }
                                                else
                                                {
                                                    if (siteNameCellResults != null)
                                                    {
                                                        new_sheet.Cell(c_remarks, 2).Value = "'Nom Global Report' vide pour '" + template_sheet.Cell(i, siteNameCellResults.GetAddress().Column).Value.ToString() + "'";
                                                    }
                                                    else
                                                    {
                                                        new_sheet.Cell(c_remarks, 2).Value = "'Nom Global Report' vide pour ligne " + i;
                                                    }
                                                    c_remarks++;
                                                }
                                                Console.WriteLine("Line " + i + " treated");
                                            }
                                        }
                                        else
                                        {
                                            new_sheet.Cell(c_remarks, 2).Value = "'Nom du contrat' not found in Matching contract-site sheet";
                                        }
                                    }
                                    else
                                    {
                                        new_sheet.Cell(c_remarks, 2).Value = "There is no Site in Results sheet";
                                    }
                                }
                                else
                                {
                                    new_sheet.Cell(c_remarks, 2).Value = "'Nom Global Report' not found in Results sheet";
                                }
                            }
                            //Otherwise, it concerns the Contracts
                            else
                            {
                                var lumworkCell = template_sheet.Find("LumWork action", false, false, false);
                                var lumworkResultCell = template_sheet.Find("LumWork action result", false, false, false);

                                if (lumworkCell == null || lumworkResultCell == null)
                                {
                                    MessageBox.Show("Columns 'LumWork action' and 'LumWork action result' not found.");
                                }
                                else
                                {
                                    var columns_to_chek = CheckColumns(document, template_sheet, data_sheet, matching_sheet);
                                    string globalReportName = GetGlobalReportName(matching_sheet, "Nom contrat Global report");

                                    int position_in_data = FindColumnByName(data_sheet, globalReportName);
                                    int position_in_template = FindColumnByName(template_sheet, "Nom contrat Global report");

                                    int nbRowsToRetrieve = GetLastEmptyRow(data_sheet, globalReportName);
                                    //int nbRowsToRetrieve = data_sheet.NotEmptyColumnMax;

                                    if (position_in_data != -1 && position_in_template != -1)
                                    {
                                        int i_col1 = 1;
                                        int i_treatment = 0;
                                        for (int i = 1; i < nbRowsToRetrieve + 1; i++)
                                        //for (int i = 1; i < nbRowsToRetrieve; i++)
                                        {
                                            var contract_name_cell = data_sheet.Cell(i, position_in_data).Value;
                                            if (contract_name_cell != null && contract_name_cell.ToString() != "")
                                            {
                                                var contract_name = contract_name_cell.ToString();

                                                int pos = ContractExists(template_sheet, contract_name, position_in_template);

                                                int last_row = template_sheet.NotEmptyRowMax;

                                                //List<string> dataColumns = GetSheetColumns(template_sheet);

                                                //List<List<string>> data = RetrieveRowData(data_sheet, matching_sheet, dataColumns, i, new_sheet, i_col1);
                                                List<List<string>> data = RetrieveRowData(data_sheet, matching_sheet, columns_to_chek, i, new_sheet, i_col1);

                                                //If record exists
                                                if (pos != -1)
                                                {
                                                    for (int j = 0; j < data.Count(); j++)
                                                    {
                                                        var column_name = data[j][0];

                                                        if (column_name != "")
                                                        {
                                                            var colValue = data[j][1];
                                                            var colAddress = template_sheet.Find(column_name, false, false, false).GetAddress();

                                                            if (column_name != "Nom du contrat")
                                                            {
                                                                template_sheet.Cell(pos, colAddress.Column).Value = colValue;
                                                            }
                                                        }
                                                    }

                                                    template_sheet.Cell(pos, lumworkCell.GetAddress().Column).Value = "Update";

                                                    //int worked = CopyRow(data_sheet, template_sheet, matching_sheet, i);

                                                    //if (worked != -1)
                                                    //{
                                                    //    template_sheet.Cell(pos, lumworkCell.GetAddress().Column).Value = "Update";
                                                    //    //treated_sheet.Cell(i - 1, 0).Value = "'" + contract_name + "' contract modified";
                                                    //}
                                                    //else
                                                    //{
                                                    //    treated_sheet.Cell(i_treatment, 0).Value = "'" + contract_name + "' contract not modified";
                                                    //    i_treatment++;
                                                    //}
                                                }
                                                else
                                                {
                                                    for (int j = 0; j < data.Count(); j++)
                                                    {
                                                        var col_name = data[j][0];

                                                        if (col_name != "")
                                                        {
                                                            var colValue = data[j][1];
                                                            var colAddress = template_sheet.Find(col_name, false, false, false).GetAddress();

                                                            template_sheet.Cell(last_row + 1, colAddress.Column).Value = colValue;
                                                        }
                                                    }
                                                    template_sheet.Cell(last_row + 1, lumworkCell.GetAddress().Column).Value = "Insert";
                                                    //treated_sheet.Cell(i - 1, 0).Value = "'" + contract_name + "' contract inserted";
                                                }
                                            }
                                            else
                                            {
                                                new_sheet.Cell(i_col1, 1).Value = "'Contract_name' empty in Data sheet, line " + i;
                                                i_col1++;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (position_in_data == -1)
                                        {
                                            //new_sheet.Cell(1, 1).Value = "'Nom Global Report' not found in Data sheet";
                                            MessageBox.Show("'Nom Global Report' not found in Data sheet");
                                        }
                                        if (position_in_template == -1)
                                        {
                                            //new_sheet.Cell(1, 3).Value = "'Nom Global Report' not found in Results sheet";
                                            MessageBox.Show("'Nom Global Report' not found in Results sheet");
                                        }
                                    }
                                }
                            }

                            Cursor.Current = Cursors.Default;

                            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

                            saveFileDialog1.Filter = "Excel |*.xlsx";
                            saveFileDialog1.FilterIndex = 2;
                            saveFileDialog1.RestoreDirectory = true;

                            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                            {
                                Cursor.Current = Cursors.WaitCursor;

                                // Save document
                                document.SaveAs(saveFileDialog1.FileName);

                                // Close Spreadsheet
                                document.Close();

                                Cursor.Current = Cursors.Default;

                                MessageBox.Show("File generated with success in : " + saveFileDialog1.FileName);

                                // open generated XLS document in default program
                                Process.Start(saveFileDialog1.FileName);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Cursor.Current = Cursors.Default;
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        //public string GetGlobalReportName(Worksheet sheet, string columnName)
        //{
        //    int lastUsedRow = sheet.NotEmptyRowMax;
        //    //int lastUsedRow = GetLastEmptyRow(sheet);
        //    string name = "";
        //    var cell_reflet = sheet.Find("Libelle reflet", false, false, false);
        //    var cell_global_report = sheet.Find("Libelle enablon", false, false, false);

        //    if (cell_reflet != null && cell_global_report != null)
        //    {
        //        for (int i = 0; i < lastUsedRow + 1; i++)
        //        {
        //            var refletName = sheet.Cell(i, cell_reflet.GetAddress().Column);
        //            if (refletName != null && refletName.ToString() == columnName && sheet.Cell(i, cell_global_report.GetAddress().Column) != null && sheet.Cell(i, cell_global_report.GetAddress().Column).ToString() != "")
        //            {
        //                name = sheet.Cell(i, cell_global_report.GetAddress().Column).ToString();
        //                break;
        //            }
        //        }
        //    }

        //    return name;
        //}

        public string GetGlobalReportName(Worksheet sheet, string columnName)
        {
            var reflet_col = sheet.Find("Libelle reflet", false, false, false);
            var gr_col = sheet.Find("Libelle enablon", false, false, false);
            string name_gr = "";

            if (reflet_col != null && gr_col != null)
            {
                var name_reflet = sheet.Find(columnName, false, false, false);

                if (name_reflet != null && sheet.Cell(name_reflet.GetAddress().Row, gr_col.GetAddress().Column) != null)
                {
                    name_gr = sheet.Cell(name_reflet.GetAddress().Row, gr_col.GetAddress().Column).ToString();
                }
            }

            return name_gr;
        }

        public int FindColumnByName(Worksheet sheet, string columnName)
        {
            int lastUsedCol = sheet.UsedRangeColumnMax;
            int position = -1;

            for (int i = 0; i < lastUsedCol + 1; i++)
            {
                if (sheet.Cell(0, i).Value != null && sheet.Cell(0, i).Value.ToString() == columnName)
                {
                    position = i;
                    break;
                }
            }

            return position;
        }

        public int ContractExists(Worksheet sheet, string contractName, int columnIndex)
        {
            int lastUsedRow = sheet.NotEmptyRowMax;
            //int lastUsedRow = GetLastEmptyRow(sheet);
            int position = -1;

            for (int i = 0; i < lastUsedRow + 1; i++)
            {
                var cell = sheet.Cell(i, columnIndex);
                if (cell.Value != null && cell.Value.ToString() == contractName)
                {
                    position = i;
                    break;
                }
            }

            return position;
        }

        //public List<string> GetSheetColumns(Worksheet sheet)
        //{
        //    int lastUsedCol = sheet.NotEmptyColumnMax;

        //    List<string> columns = new List<string>();

        //    for (int i = 0; i < lastUsedCol + 1; i++)
        //    {
        //        var columnName = sheet.Cell(0, i).ToString();
        //        if (columnName != "LumWork form ID" && columnName != "LumWork version ID" && columnName != "LumWork link" && columnName != "LumWork action" && columnName != "LumWork action result")
        //        {
        //            columns.Add(columnName);
        //        }
        //    }

        //    return columns;
        //}

        public List<string> GetSheetColumns(Worksheet sheet)
        {
            List<string> columns = new List<string>();
            var current_col = 0;

            while (sheet.Cell(0, current_col).Value != null && sheet.Cell(0, current_col).Value.ToString() != "")
            {
                var col_name = sheet.Cell(0, current_col).Value.ToString();
                if (col_name != "LumWork form ID" && col_name != "LumWork version ID" && col_name != "LumWork link" && col_name != "LumWork action" && col_name != "LumWork action result")
                {
                    columns.Add(col_name);
                }
                current_col++;
            }

            return columns;
        }

        //public struct Column
        //{
        //    public int Id;
        //    public string Name;
        //    public string GrName;
        //    public string Value;
        //    public int Coeff;
        //}

        public List<string> CheckColumns(Spreadsheet document, Worksheet source, Worksheet dest, Worksheet matching)
        {
            Worksheet new_logs = document.Workbook.Worksheets.Add("Logs - Matching");
            //Initialization of "Logs - Matching" sheet
            new_logs.Cell(0, 0).Value = "Columns not found in matching sheet";
            int row = 1;

            var columns_source = GetSheetColumns(source);
            var columns_to_retrieve = new List<string>();

            foreach(var col in columns_source)
            {
                var gr_name = GetGlobalReportName(matching, col);

                if (gr_name != "")
                {
                    //columns_to_retrieve.Add(gr_name);
                    columns_to_retrieve.Add(col);
                }
                else
                {
                    new_logs.Cell(row, 0).Value = col;
                    row++;
                }
            }

            return columns_to_retrieve;
        }

        public List<List<string>> RetrieveRowData(Worksheet sheet, Worksheet matchSheet, List<string> columns, int rowNumber, Worksheet logSheet, int last_row_col1)
        {
            List<List<string>> data = new List<List<string>>();

            int i_col0 = 1;
            int i_col1 = last_row_col1 + 1;

            foreach (var col in columns)
            {
                List<string> rowData = new List<string>();

                var globRepName = GetGlobalReportName(matchSheet, col);
                var cell = matchSheet.Find(col, false, false, false);

                if (cell != null)
                {
                    var address = cell.GetAddress();
                    var row = address.Row;
                    var colNumber = FindColumnByName(sheet, globRepName);
                    if (colNumber != -1 && sheet.Cell(rowNumber, colNumber) != null)
                    {
                        var cell_coeff = matchSheet.Find("Coefficient", false, false, false);
                        if (cell_coeff != null)
                        {
                            if (matchSheet.Cell(row, cell_coeff.GetAddress().Column).Value != null)
                            {
                                var coeff = matchSheet.Cell(row, cell_coeff.GetAddress().Column).ToInteger();
                                rowData.Add(col);
                                if (sheet.Cell(rowNumber, colNumber).ToString() != "")
                                {
                                    var value = double.Parse(sheet.Cell(rowNumber, colNumber).ToString());
                                    var sum = value * coeff;
                                    rowData.Add(sum.ToString());
                                }
                                else
                                {
                                    rowData.Add("");
                                }
                            }
                            else
                            {
                                rowData.Add(col);
                                rowData.Add(sheet.Cell(rowNumber, colNumber).ToString());
                            }
                        }
                    }
                    else
                    {
                        var colInLog = logSheet.Find(col, false, false, false);

                        if (colInLog == null)
                        {
                            //int last_row_logs = logSheet.NotEmptyRowMax;
                            //int last_row_logs = GetLastEmptyRow(logSheet, "From data");
                            //logSheet.Cell(last_row_logs + 1, 1).Value = "Column '" + col + "' not found in data sheet";
                            logSheet.Cell(i_col1, 0).Value = "Column '" + col + "' not found in data sheet";
                            i_col1++;
                        }

                        rowData.Add("");
                    }
                    data.Add(rowData);
                }
                else
                {
                    var colInLog = logSheet.Find(col, false, false, false);

                    if (colInLog == null)
                    {
                        //int last_row_logs = logSheet.NotEmptyRowMax;
                        //int last_row_logs = GetLastEmptyRow(logSheet, "From matching");
                        //logSheet.Cell(last_row_logs + 1, 0).Value = "Column '" + col + "' not found in matching sheet";
                        logSheet.Cell(i_col0, 0).Value = "Column '" + col + "' not found in matching sheet";
                        i_col0++;
                    }
                }
            }

            return data;
        }

        public int GetLastEmptyRow(Worksheet sheet, string idColumn)
        {
            int index = -1;
            var column = sheet.Find(idColumn, false, false, false);

            if (column != null)
            {
                int i = 0;
                while (sheet.Cell(i, column.GetAddress().Column).Value != null && sheet.Cell(i, column.GetAddress().Column).Value.ToString() != "")
                {
                    i++;
                }
                index = i;
            }

            return index;
        }

        public string GetRefletName(Worksheet sheet, string columnName)
        {
            string translation = "";
            var grCol = sheet.Find("Libelle enablon", false, false, false);

            if (grCol != null)
            {
                var refletCol = sheet.Find("Libelle reflet", false, false, false);

                if (refletCol != null)
                {
                    var searchedCol = sheet.Find(columnName, false, false, false);

                    if (searchedCol != null)
                    {
                        translation = sheet.Cell(searchedCol.GetAddress().Row, refletCol.GetAddress().Column).Value.ToString();
                    }
                    else
                    {

                    }
                }
                else
                {

                }
            }
            else
            {

            }
            return translation;
        }

        public int CopyRow(Worksheet source_sheet, Worksheet destination_sheet, Worksheet matching_sheet, int row)
        {
            int returnCode = -1;
            var idCellSource = source_sheet.Find("Global_report_name", false, false, false);

            //If Global_report_name column exists in source_sheet
            if (idCellSource != null)
            {
                var idCellDest = destination_sheet.Find("Nom contrat Global report", false, false, false);

                //If Nom contrat Global report column exists in destination_sheet
                if (idCellDest != null)
                {
                    var gr_id = source_sheet.Cell(row, idCellSource.GetAddress().Column).Value.ToString();

                    //If Global report code in source_sheet is not empty
                    if (gr_id != "")
                    {
                        //If exists, update
                        if (destination_sheet.Find(gr_id, false, false, false) != null)
                        {
                            returnCode = 0;
                        }
                        //Add new line
                        else
                        {
                            var last_used_row = destination_sheet.NotEmptyRowMax;
                            destination_sheet.Cell(last_used_row + 1, idCellDest.GetAddress().Column).Value = source_sheet.Cell(row, idCellSource.GetAddress().Column).Value;
                            returnCode = 0;
                        }
                        //Loop all data columns except Contract_name && Global_report_name
                        //for (int i = 0; i < source_sheet.NotEmptyColumnMax; i++)
                        //{
                        //    if (source_sheet.Cell(0, i).Value != null && source_sheet.Cell(0, i).Value.ToString() != "")
                        //    {
                        //        var col_name = source_sheet.Cell(0, i).Value.ToString();
                        //        if (col_name != "Contract_name" && col_name != "Global_report_name")
                        //        {
                        //            var translated_col = GetRefletName(matching_sheet, col_name);

                        //            //If we can get the translation of the column from the matching_sheet
                        //            if (translated_col != "")
                        //            {
                        //                var translated_col_dest = destination_sheet.Find(translated_col, false, false, false);

                        //                //If translated column exists in destination_sheet
                        //                if (translated_col_dest != null)
                        //                {
                        //                    var row_in_dest = destination_sheet.Find(gr_id, false, false, false);

                        //                    //If the Global report code exists in destination_sheet, update the row
                        //                    if (row_in_dest != null)
                        //                    {
                        //                        var row_dest = row_in_dest.GetAddress().Row;
                        //                        var col_dest = translated_col_dest.GetAddress().Column;
                        //                        destination_sheet.Cell(row_dest, col_dest).Value = source_sheet.Cell(row, i).Value;
                        //                        returnCode = 0;
                        //                    }
                        //                    //If the Global report code does not exists in destination_sheet, insert the row
                        //                    else
                        //                    {
                        //                        var last_used_row = destination_sheet.NotEmptyRowMax;
                        //                        destination_sheet.Cell(last_used_row + 1, translated_col_dest.GetAddress().Column).Value = source_sheet.Cell(row, i).Value;
                        //                        returnCode = 1;
                        //                    }
                        //                }
                        //                else
                        //                {

                        //                }
                        //            }
                        //        }
                        //    }
                        //}
                    }
                }
                else
                {
                    //MessageBox.Show("<Nom contrat Global report> column not found in " + destination_sheet.Name);
                }
            }
            else
            {
                //MessageBox.Show("<Global_report_name> column not found in " + source_sheet.Name);
            }
            return returnCode;
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void openExempleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //string path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Example\Example.xlsx");
            //Assembly _assembly;
            //StreamReader _textStreamReader;

            try
            {
                string resourceName = "Example.xlsx";
                string filename = Path.Combine(Path.GetTempPath(), resourceName);

                Assembly asm = typeof(Program).Assembly;
                using (Stream stream = asm.GetManifestResourceStream(
                    asm.GetName().Name + "." + resourceName))
                {
                    using (Stream output = new FileStream(filename,
                        FileMode.OpenOrCreate, FileAccess.Write))
                    {
                        byte[] buffer = new byte[32 * 1024];
                        int read;
                        while ((read = stream.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            output.Write(buffer, 0, read);
                        }
                    }
                }
                Process.Start(filename);
            }
            catch
            {
                MessageBox.Show("Error accessing resources!");
            }
            //Process.Start(path);
        }
    }
}
