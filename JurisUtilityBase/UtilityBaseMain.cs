﻿using System;
using System.Linq;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using JDataEngine;
using JurisAuthenticator;
using System.Data.SqlClient;
using JurisSVR;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Text.RegularExpressions;
using JurisWrapper;



namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {


        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        private string path = "";

        JurisUtility _jurisUtility;

        List<string> errors = new List<string>();
        List<Bill> badBills = new List<Bill>();



        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
            //            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
            //            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }
            comboBox1.Items.Clear();
            //only list clients with invoices
            string SQLPC2 = "select dbo.jfn_FormatClientCode(clicode)  + '    ' + clireportingname as PC from client " +
                            " where clisysnbr in (select distinct clisysnbr " +
                            " from ledgerhistory inner join arbill on arbillnbr = lhbillnbr  " +
                            " inner join matter on matsysnbr = lhmatter inner join client on clisysnbr = matclinbr  " +
                            " where lhtype in ('3', '4') and lhbillnbr not in (select lhbillnbr from ledgerhistory where lhtype in ('A', 'B', 'C'))) " +
                            " order by dbo.jfn_FormatClientCode(clicode)";
            DataSet myRSPC2 = _jurisUtility.RecordsetFromSQL(SQLPC2);

            if (myRSPC2 == null || myRSPC2.Tables.Count == 0 || myRSPC2.Tables[0].Rows.Count == 0)
                MessageBox.Show("There are no Clients. Correct and run the tool again", "Client Error");
            else
            {
                foreach (DataRow dr in myRSPC2.Tables[0].Rows)
                    comboBox1.Items.Add(dr["PC"].ToString());
                comboBox1.SelectedIndex = 0;
            }



        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {

        }


        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum;
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>


        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }



        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {

            errors.Clear();
            badBills.Clear();
            richTextBox1.Text = "";
            if (verifyBoxes())
            {
                UpdateStatus("Gathering Information...", 0, 3);
                List<Bill> bills = getBills().ToList();
                int total = bills.Count;
                int runningTotal = 1;
                foreach (Bill bb1 in bills)
                {
                    if (!bb1.badBill)
                    {
                        try
                        {
                            string finalPath = Path.Combine(path, getFileName(bb1));

                            GetBillImage1(bb1.billNo.ToString(), finalPath);
                            if (bb1.hasExpAttach && checkBoxExpense.Checked)
                            {
                                foreach (ExpAttachment eex in bb1.exps)
                                {
                                    string expPath = finalPath.Replace(".pdf", "");
                                    expPath = expPath + "-" + eex.fileName;
                                    File.WriteAllBytes(expPath, eex.fileData);
                                }
                            }

                        }
                        catch (Exception ccs) { errors.Add(bb1.billNo.ToString() + " encounter an error: " + ccs.Message); }
                    }
                    UpdateStatus("Converting Invoices...", runningTotal, total);
                    runningTotal++;
                }


                UpdateStatus("Process Complete!", total, total);
                foreach (Bill bb2 in badBills)
                    errors.Add("There is no archive image for bill number " + bb2.billNo);
                if (errors.Count == 0)
                {
                    DialogResult dr = MessageBox.Show("Process completed without error. Do you want to view the files?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dr == DialogResult.Yes)
                        Process.Start(path);
                }
                else
                {
                    DialogResult dr = MessageBox.Show("Process completed but there were problems with some" + "\r\n" + "invoices. Do you want to view the error log?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dr == DialogResult.Yes)
                    {
                        foreach (string ss in errors)
                            richTextBox1.Text = richTextBox1.Text + ss + "\r\n";
                        tabControl1.SelectedIndex = 2;
                    }

                }

            }
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {

            this.Close();

        }

        private void checkBoxByDate_CheckedChanged(object sender, EventArgs e)
        {
            textBoxDate1.Visible = checkBoxByDate.Checked;
            textBoxDate2.Visible = checkBoxByDate.Checked;
            labelDate.Visible = checkBoxByDate.Checked;
            if (checkBoxByDate.Checked || checkBoxByNumber.Checked || checkBoxByCliMat.Checked)
                button1.Enabled = true;
            else
                button1.Enabled = false;
        }

        private void checkBoxByNumber_CheckedChanged(object sender, EventArgs e)
        {
            textBoxNum1.Visible = checkBoxByNumber.Checked;
            textBoxNum2.Visible = checkBoxByNumber.Checked;
            labelNum.Visible = checkBoxByNumber.Checked;
            if (checkBoxByDate.Checked || checkBoxByNumber.Checked || checkBoxByCliMat.Checked)
                button1.Enabled = true;
            else
                button1.Enabled = false;
        }

        private void checkBoxByCliMat_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxByDate.Checked || checkBoxByNumber.Checked || checkBoxByCliMat.Checked)
                button1.Enabled = true;
            else
                button1.Enabled = false;
            comboBox1.Visible = checkBoxByCliMat.Checked;

        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowNewFolderButton = true;
            // Show the FolderBrowserDialog.  
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                label1.Text = folderBrowserDialog1.SelectedPath;
                path = folderBrowserDialog1.SelectedPath;
                checkBoxByCliMat.Enabled = true;
                checkBoxByDate.Enabled = true;
                checkBoxByNumber.Enabled = true;
                label2.Enabled = true;
                textBoxNaming.Enabled = true;
                checkBoxBalance.Enabled = true;
            }
        }

        private List<Bill> getBills()
        {
            List<Bill> bills = new List<Bill>();
            string sql = "";

            if (checkBoxBalance.Checked)
            {
                sql = "select distinct matsysnbr, clisysnbr,dbo.jfn_FormatClientCode(clicode) as clicode, clireportingname, " +
                        " dbo.jfn_FormatMatterCode(MatCode) as matcode, matreportingname,arbillnbr, convert(varchar, arbilldate, 101) as BillDate " +
                        " from armatalloc inner join arbill on arbillnbr = armbillnbr " +
                        " inner join matter on matsysnbr = armmatter inner join client on clisysnbr = matclinbr " +
                        " where ARMBalDue <> 0 ";

            }
            else
            {
                sql = "select distinct matsysnbr, clisysnbr,dbo.jfn_FormatClientCode(clicode) as clicode, clireportingname, " +
                        " dbo.jfn_FormatMatterCode(MatCode) as matcode, matreportingname,arbillnbr, convert(varchar, arbilldate, 101) as BillDate " +
                        " from ledgerhistory inner join arbill on arbillnbr = lhbillnbr " +
                        " inner join matter on matsysnbr = lhmatter inner join client on clisysnbr = matclinbr " +
                        " where lhtype in ('3', '4') and lhbillnbr not in (select lhbillnbr from ledgerhistory where lhtype in ('A', 'B', 'C')) ";
            }

            List<string> whereClauses = new List<string>();
            if (checkBoxByDate.Checked)
                whereClauses.Add(" arbilldate between '" + textBoxDate1.Text + "' and '" + textBoxDate2.Text + "' ");
            if (checkBoxByNumber.Checked)
                whereClauses.Add(" arbillnbr between " + textBoxNum1.Text + " and " + textBoxNum2.Text + " ");
            if (checkBoxByCliMat.Checked)
                whereClauses.Add(" dbo.jfn_FormatClientCode(clicode) = '" + this.comboBox1.GetItemText(this.comboBox1.SelectedItem).Split(' ')[0] + "' ");

            string endingWhere = "";



            if (whereClauses.Count > 0)
            {
                endingWhere = " and ";
                foreach (string clause in whereClauses)
                    endingWhere = endingWhere + clause + " and";
            }
            endingWhere = endingWhere.Substring(0, endingWhere.Length - "and".Length);
            sql = sql + endingWhere + " order by arbillnbr";

            UpdateStatus("Gathering Information...", 1, 3);
            Bill bill = null;
            //now get the bills
            DataSet ds = _jurisUtility.RecordsetFromSQL(sql);
            if (ds == null || ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0)
                MessageBox.Show("Your selections returns no invoices. Please refine them so invoices are included", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                JBillsUtility jbills = new JBillsUtility();
                jbills.SetInstance(CompanyCode);
                JurisDbName = jbills.Company.DatabaseName;
                JBillsDbName = "JBills" + jbills.Company.Code;
                jbills.OpenDatabase();
                if (jbills.DbOpen)
                {
                    ///GetFieldLengths();
                }
                string sql1 = "select distinct BASInvoiceNbr from BillArchiveSegment where BASInvoiceNbr <> 0";
                DataSet rs = jbills.RecordsetFromSQL(sql1);
                if (rs == null || rs.Tables.Count == 0 || rs.Tables[0].Rows.Count == 0)
                    MessageBox.Show("This Juris database has no archive bill images", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {

                        bill = new Bill();
                        bill.matsys = Convert.ToInt32(dr["matsysnbr"].ToString());
                        bill.clisys = Convert.ToInt32(dr["clisysnbr"].ToString());
                        bill.matterNo = dr["matcode"].ToString();
                        bill.clientNo = dr["clicode"].ToString();
                        bill.matterName = dr["matreportingname"].ToString();
                        bill.clientName = dr["clireportingname"].ToString();
                        bill.billNo = Convert.ToInt32(dr["arbillnbr"].ToString());
                        bill.billDate = dr["BillDate"].ToString();
                        bill.badBill = true;
                        bill.hasExpAttach = false;
                        bills.Add(bill);

                    }
                    ds.Clear();
                    UpdateStatus("Gathering Information...", 2, 3);
                    //now that we have all the bills, see which ones arent valid and mark them
                    foreach (Bill bb in bills)
                    {
                        foreach (DataRow dr in rs.Tables[0].Rows)
                        {
                            if (bb.billNo == Convert.ToInt32(dr["BASInvoiceNbr"].ToString()))
                            {
                                bb.badBill = false;
                                break;
                            }
                        }
                    }
                    jbills.CloseDatabase();

                    string sql2 = "  select name, attachmentobject, bebillnbr from Attachment aa " +
                                  "  inner join ExpenseEntryAttachment ea on ea.AttachmentID = aa.id " +
                                  "  inner join expenseEntry ee on ee.entryid = ea.EntryID " +
                                  "  INNER JOIN  ExpenseEntryLink el ON ee.EntryID = el.EntryID " +
                                  "  INNER JOIN BilledExpenses be ON el.EBDID = be.beid " +
                                  "  where AttachmentType = 0";
                    DataSet fs = _jurisUtility.RecordsetFromSQL(sql2);
                    foreach (DataRow dr in fs.Tables[0].Rows)
                    {
                        foreach (Bill bb in bills)
                        {
                            if (bb.billNo == Convert.ToInt32(dr["bebillnbr"].ToString()))
                            {
                                bb.hasExpAttach = true;
                                ExpAttachment expattach = new ExpAttachment();
                                expattach.fileName = dr["name"].ToString();
                                expattach.fileData = (byte[])dr["attachmentobject"];
                                bb.exps.Add(expattach);
                                break;
                            }
                        }
                    }

                    fs.Clear();

                    _jurisUtility.CloseDatabase();
                    foreach (Bill ba in bills)
                    {
                        if (ba.badBill)
                        {
                            badBills.Add(ba);
                        }
                    }

                }


            }
            UpdateStatus("Gathering Information...", 3, 3);
            return bills;
        }

        private bool verifyBoxes()
        {

            if (checkBoxByDate.Checked)
            {
                if (!verifyDate(textBoxDate1.Text) || !verifyDate(textBoxDate2.Text))
                {
                    MessageBox.Show("Both date boxes must be filled in and have a valid date if selecting by Date", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            if (checkBoxByNumber.Checked)
            {
                if (!verifyNum(textBoxNum1.Text) || !verifyNum(textBoxNum2.Text))
                {
                    MessageBox.Show("Both Number boxes must be filled in and have a valid integer if selecting by Bill Number", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            if (!verifyName())
            {
                MessageBox.Show("The naming convention is incorrect", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }

        private bool verifyDate(string dt)
        {
            try
            {
                Convert.ToDateTime(dt);
                return true;
            }
            catch (Exception ee)
            {
                return false;
            }
        }

        private bool verifyNum(string dt)
        {
            try
            {
                Convert.ToInt32(dt);
                return true;
            }
            catch (Exception ee)
            {
                return false;
            }
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step / steps) * 100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
            this.Refresh();
        }

        private string getFileName(Bill bb)
        {
            string output = textBoxNaming.Text.Replace(".pdf", "");
            output = output.Replace("pdf", "");

            output = output.Replace("[ClientNum]", bb.clientNo);
            output = output.Replace("[MatterNum]", bb.matterNo);
            output = output.Replace("[ClientName]", bb.clientName);
            output = output.Replace("[MatterName]", bb.matterName);
            output = output.Replace("[BillNum]", bb.billNo.ToString());
            output = output.Replace("[BillDate]", bb.billDate.Replace("/", "-"));
            output = output.Replace("[Clisys]", bb.clisys.ToString());
            output = output.Replace("[Matsys]", bb.matsys.ToString());
            output = output.Replace("[NowDate]", DateTime.Now.ToString("MM/dd/yyyy").Replace("/", "-"));
            output = output.Replace("[NowTime]", DateTime.Now.ToString("MM/dd/yyyy hh:mm tt").Replace("/", "-").Replace(":", " "));
            output = output + ".pdf";
            return output;
        }

        private bool verifyName()
        {
            string text = textBoxNaming.Text.Trim();
            text = text.ToLower();
            if (!text.EndsWith(".pdf"))
                return false;
            else
            {
                text = text.Replace("[clientnum]", "").Replace("[matternum]", "").Replace("[clientname]", "").Replace("[mattername]", "").Replace("[billnum]", "")
                    .Replace("[billdate]", "").Replace("[clisys]", "").Replace("[matsys]", "").Replace("[nowdate]", "").Replace("[nowtime]", "").Replace(".pdf", "")
                    .Replace("-", "").Replace(" ", "");
                if (!string.IsNullOrEmpty(text))
                    return false;
                else
                    return true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(richTextBox1.Text);
        }

        private static Wrapper _jurisWrapper = null;

        private Wrapper JurisWrapper
        {
            get { return _jurisWrapper ?? (_jurisWrapper = new Wrapper()); }
        }

        public void GetBillImage2(string billNumber, string finalPath)
        {

            try
            {
                Wrapper wrapper = new Wrapper();
                if (!wrapper.LogonCompany(CompanyCode.Replace("Company", "")))
                {
                    MessageBox.Show("Unable to view bill image.  Please verify that the company is properly licensed. ");
                    this.Close();
                }
                else
                {

                    try
                    {
                        var result = wrapper.GetBillImage(Convert.ToInt32(billNumber), finalPath);
                        Cursor = Cursors.Default;

                        if (result == -1)
                        {
                            var errorMessage = String.Format("Unable to open archive image for bill number {0}.",
                                billNumber);
                            errors.Add(errorMessage);
                        }
                        else if (result == 9)
                        {
                            errors.Add("There is no archive image for bill number");
                        }
                        else
                        {

                        }

                        if (Marshal.IsComObject(wrapper))
                        {
                            while (Marshal.ReleaseComObject(wrapper) > 0) ;
                        }
                    }
                    catch (Exception exception)
                    {
                        errors.Add(exception.Message + " : " + billNumber + " : " + finalPath);
                    }
                }
            }
            catch (Exception exc)
            {
                errors.Add(exc.Message + " : " + billNumber + " : " + finalPath);

            }
        }



        public void GetBillImage1(string billNumber, string finalPath)
        {
            try
            {
                Wrapper wrapper = null;
                try
                {
                    wrapper = JurisWrapper;
                }
                catch (Exception exception)
                {
                    var errorMessage = exception.Message;
                    if (wrapper != null && wrapper.WrapperException != null)
                    {
                        errorMessage += " " + wrapper.WrapperException.Message;
                    }
                    MessageBox.Show("Unable to view bill image...null wrapper: " + errorMessage,
    Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                }
                if (wrapper == null)
                {
                    MessageBox.Show("Unable to view bill image...null wrapper",
                        Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                    Cursor = Cursors.Default;
                }
                else if (!wrapper.LogonCompany(CompanyCode.Replace("Company", "")))
                {
                    MessageBox.Show(string.Format("Unable to view bill image.  Please verify that the company is properly licensed. {0}", wrapper.WrapperException != null ? wrapper.WrapperException.Message : string.Empty),
                        Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                    Cursor = Cursors.Default;
                }
                else
                {


                    try
                    {
                        var result = wrapper.GetBillImage(Convert.ToInt32(billNumber), finalPath);
                        Cursor = Cursors.Default;

                        if (result == -1)
                        {
                            var errorMessage = String.Format("Unable to open archive image for bill number {0}.",
                                billNumber);
                            if (wrapper.WrapperException != null)
                            {
                                errorMessage += " " + wrapper.WrapperException.Message;
                            }
                            errors.Add(errorMessage);
                        }
                        else if (result == 9)
                        {
                            errors.Add(String.Format("There is no archive image for bill number {0}", billNumber));
                        }
                        else
                        {
                            //var proc = new Process { EnableRaisingEvents = false, StartInfo = { FileName = path } };
                            //proc.Start();
                        }

                        if (Marshal.IsComObject(wrapper))
                        {
                            while (Marshal.ReleaseComObject(wrapper) > 0) ;
                        }
                    }
                    catch (Exception exception)
                    {
                        var errorMessage = exception.Message;
                        if (wrapper.WrapperException != null)
                        {
                            errorMessage += " " + wrapper.WrapperException.Message;

                            if (errorMessage.IndexOf("juriswrapper.log", StringComparison.OrdinalIgnoreCase) >= 0)
                                errors.Add("Bill Number: " + billNumber + " appears to be in an incorrect format and may need to be repaired");
                            else
                                errors.Add("Bill Number: " + billNumber + " could not be accessed because " + errorMessage);
                            Cursor = Cursors.Default;
                        }
                    }
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Error retrieving bill image outer catch", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Cursor = Cursors.Default;
            }
        }







        private void checkBoxExpense_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}