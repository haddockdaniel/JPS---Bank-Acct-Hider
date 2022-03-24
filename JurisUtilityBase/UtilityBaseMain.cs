using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public int FldClient { get; set; }

        public int FldMatter { get; set; }

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

            comboBox1.ClearItems();
            string SQLPC2 = "select BnkCode  + '    ' + left(BnkDesc, 30) as PC from BankAccount order by BnkCode";
            DataSet myRSPC2 = _jurisUtility.RecordsetFromSQL(SQLPC2);

            if (myRSPC2 == null || myRSPC2.Tables.Count == 0 || myRSPC2.Tables[0].Rows.Count == 0)
            {
                MessageBox.Show("Juris has no Bank Accounts set up. The tool will not exit", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                this.Close();
            }
            else
            {
                foreach (DataRow dr in myRSPC2.Tables[0].Rows)
                    comboBox1.Items.Add(dr["PC"].ToString());
                comboBox1.SelectedIndex = 0;
            }

        }



        #endregion

        #region Private methods
        //this.comboBox1.GetItemText(this.comboBox1.SelectedItem).Split(' ')[0]
        private void DoDaFix()
        {
            bool off = false;
            bool time = false;

            //see if its assigned to an office
            string sql = "  select bankaccount.BnkCode, case when OGAOffice is null then 'Not There' else OGAOffice end as Office from bankaccount " +
                          " inner join BkAcctGLAcct on BGABkAcct = bnkcode and BGAType = 'A' " +
                          " left outer join OfficeGLAccount on OGAAcct = BGAAcct " +
                          " where bnkcode = '" + this.comboBox1.GetItemText(this.comboBox1.SelectedItem).Split(' ')[0] + "'"; 

                DataSet myRSPC2 = _jurisUtility.RecordsetFromSQL(sql);
            if (myRSPC2 == null || myRSPC2.Tables.Count == 0 || myRSPC2.Tables[0].Rows.Count == 0) // it was never assigned to a gl account
            {

            }
            else
            {
                foreach (DataRow dr in myRSPC2.Tables[0].Rows)
                {
                    if (!dr[1].ToString().Equals("Not There"))
                    {
                        MessageBox.Show("That Bank Account is tied to an office (" + dr[1].ToString() + "). Please remove it before moving forward", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        break;
                    }
                    else
                        off = true;
                }
            }
            myRSPC2.Clear();

            //any activity in checkbook in last 9 months
            sql = "  SELECT max([CkRegDate]) as DT, CkRegBank " +
                  "  FROM [CheckRegister] " +
                  "  where CkRegBank = '" + this.comboBox1.GetItemText(this.comboBox1.SelectedItem).Split(' ')[0] + "' " +
                  "  group by CkRegBank " +
                  "  having max([CkRegDate]) >= DATEADD(mm, -9, GETDATE())";
            myRSPC2 = _jurisUtility.RecordsetFromSQL(sql);
            if (myRSPC2 == null || myRSPC2.Tables.Count == 0 || myRSPC2.Tables[0].Rows.Count == 0) // it has not been used in the last 9 months
            {
                time = true;
            }
            else
            {
                MessageBox.Show("That Bank Account has had activity in the last 9 months. Last transaction: " + Convert.ToDateTime(myRSPC2.Tables[0].Rows[0][0].ToString()).ToShortDateString(), "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }

            if (time && off)
            {
                DialogResult dd = MessageBox.Show("Everything is ready to delete Bank Account " + this.comboBox1.GetItemText(this.comboBox1.SelectedItem).Split(' ')[0] + "." + "\r\n" + "This cannot be undone. Proceed?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dd == DialogResult.Yes)
                {
                    sql = "  delete FROM [DocumentTree] where dtdocclass = 6400 and dtparentid = 10 and DTKeyT = '" + this.comboBox1.GetItemText(this.comboBox1.SelectedItem).Split(' ')[0] + "'";
                    _jurisUtility.ExecuteNonQuery(0, sql);
                    MessageBox.Show("The process is complete!", "Finished", MessageBoxButtons.OK, MessageBoxIcon.None);
                }
            }
            off = false;
            time = false;
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
            if (File.Exists(filePathName ))
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
            if (checkBox1.Checked)
                DoDaFix();
            else
                MessageBox.Show("Please read and check the box ensuring the bank account can be removed", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Stop);
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {

            System.Environment.Exit(0);
          
        }




    }
}
