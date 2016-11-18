#region References
using System;
using System.Windows.Forms;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data;
using System.ComponentModel;
using System.Globalization;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
#endregion

namespace Utilities
{
    #region BPCS & SQL Connections
    public class FnP
    {
        #region Variable Declarations
        static OleDbConnection BPCSconn = new OleDbConnection();
        static SqlConnection SQLConn = new SqlConnection();

        private string dbname;
        private string dbuser;
        private string dbpass;
        #endregion

        #region Constructors
        /// <summary>
        /// Creates a new Instance of FnP Class.
        /// </summary>
        public FnP()
        {
            dbname = "DB_Produccion";
            dbuser = "VisualSQL";
            dbpass = "osc13arc";
        }
        /// <summary>
        /// Creates a new Instance of FnP Class.
        /// </summary>
        /// <param name="DB_Name">Name of the SQL DataBase to connect with.</param>
        public FnP(string DB_Name)
        {
            dbname = DB_Name;
            dbuser = "VisualSQL";
            dbpass = "osc13arc";
        }
        /// <summary>
        /// Creates a new Instance of FnP Class.
        /// </summary>
        /// <param name="DB_Name">Name of the SQL DataBase to connect with.</param>
        /// <param name="DB_User">UserName of the Database to connect with.</param>
        /// <param name="DB_Pwd">UserName's Password of the Database to connect with.</param>
        public FnP(string DB_Name, string DB_User, string DB_Pwd)
        {
            dbname = DB_Name;
            dbuser = DB_User;
            dbpass = DB_Pwd;
        }
        #endregion

        #region BPCS
        /// <summary>
        ///     Open a connection to BPCS DataBase with userodbc.
        /// </summary>
        /// <returns>
        ///     Returns a string value with the status or error of the connection.
        /// </returns>
        static String OpenBPCSConnection()
        {
            BPCSconn.ConnectionString = "Provider=IBMDA400;Data Source=S10FEFB3;User Id=userodbc;Password=saq80i;Default Collection=LXF;";
            try
            {
                BPCSconn.Open();
                return "OK";
            }
            catch (Exception e)
            {
                if (BPCSconn.State != System.Data.ConnectionState.Closed)
                {
                    BPCSconn.Close();
                }
                return "Error[OpenBPCSConnection]: " + e.Message;
            }
        }

        /// <summary>
        ///     Close the existing BPCS connection.
        /// </summary>
        static void CloseBPCSConection()
        {
            if (BPCSconn.State != System.Data.ConnectionState.Closed)
            {
                BPCSconn.Close();
            }
        }

        /// <summary>
        ///     Gets the data from BPCS with the provided SQL string.
        /// </summary>
        /// <param name="sql">
        ///     String value containing the SQL string query to retrieve data from BPCS.
        /// </param>
        /// <returns>
        ///     Returns a DataTable containing the data requested through SQL query.
        /// </returns>
        public static System.Data.DataTable GetBPCSData(string sql)
        {
            System.Data.OleDb.OleDbDataAdapter MyAdapter = new OleDbDataAdapter();
            System.Data.DataTable DTBX = new System.Data.DataTable();
            OleDbCommand BPCScom = new OleDbCommand();
            string st;
            st = OpenBPCSConnection();
            if (st == "OK")
            {
                try
                {
                    BPCScom.CommandText = sql;
                    BPCScom.Connection = BPCSconn;
                    MyAdapter.SelectCommand = BPCScom;
                    MyAdapter.Fill(DTBX);
                    DTBX.TableName = "DTBX"
                        ;
                }
                catch (Exception e)
                {
                    DTBX = ErrorMsg("Error [GetBPCSData]: " + e.Message);
                }
            }
            else
            {
                DTBX = ErrorMsg(st);
            }

            CloseBPCSConection();
            return DTBX;
        }

        public static DataTable GetBPCSBOM(string Item, string EffectiveDate, string Facility, int lvl)
        {
            DataTable X = new DataTable();
            DataTable Y = new DataTable();
            int level = lvl + 1;
            string sqls;

            sqls = "SELECT ";
            sqls += "   " + Convert.ToString(level) + ", ";
            sqls += "   LTRIM(RTRIM(B.BPROD))  AS PARENT, ";
            sqls += "   LTRIM(RTRIM(B.BCHLD))  AS ITEM, ";
            sqls += "   LTRIM(RTRIM(I.IDESC))  AS DESC, ";
            sqls += "   LTRIM(RTRIM(I.IDSCE))  AS EXT_DESC, ";
            sqls += "   LTRIM(RTRIM(I.IITYP))  AS TYPE, ";
            sqls += "   LTRIM(RTRIM(B.BSEQ))   AS SEQ, ";
            sqls += "   LTRIM(RTRIM(B.BQREQ))  AS QTY_REQ, ";
            sqls += "   LTRIM(RTRIM(I.ICLAS))  AS ITEM_CLASS, ";
            sqls += "   LTRIM(RTRIM(I.IUMS))   AS STOCK_UOM, ";
            sqls += "   LTRIM(RTRIM(B.BMSCP))  AS SCRAP, ";
            sqls += "   LTRIM(RTRIM(B.BBUBB))  AS BUBBLE_NUMBER, ";
            sqls += "   LTRIM(RTrim(B.BMBOMM)) AS ALTERNATIVE, ";
            sqls += "   LTRIM(RTrim(I.IVEND))  AS VENDOR_NUM, ";
            sqls += "   LTRIM(RTrim(X.VNDNAM)) AS VENDOR_NAME, ";
            sqls += "   LTRIM(RTrim(C.CFTLVL)) AS ACT_COST, ";
            sqls += "   LTRIM(RTrim(S.CFTLVL)) AS STD_COST ";
            sqls += "FROM MBML01 AS B ";
            sqls += "   INNER JOIN IIML02 AS I ON I.IPROD = B.BCHLD ";
            sqls += "   LEFT  JOIN AVML01 AS X ON I.IVEND = X.VENDOR ";
            sqls += "   LEFT  JOIN CMFL01 AS C ON I.IPROD = C.CFPROD AND C.CFCSET = 1 AND C.CFCBKT = 1 AND C.CFFAC = '" + Facility + "' ";
            sqls += "   LEFT JOIN CMFL01 AS S ON I.IPROD = S.CFPROD AND S.CFCSET = 2 AND S.CFCBKT = 1 AND S.CFFAC = '" + Facility + "' ";
            sqls += "WHERE B.BPROD IN('" + Item + "') ";
            sqls += "   AND BMWHS IN('" + Facility + "') ";
            sqls += "   AND B.BDDIS >= " + EffectiveDate + " AND B.BDEFF <= " + EffectiveDate + " AND B.BMBOMM <> 'AL' ";
            sqls += "ORDER BY B.BPROD, B.BMBOMM, B.BSEQ ASC";

            X = GetBPCSData(sqls);

            foreach (DataColumn c in X.Columns)
            {
                Y.Columns.Add(c.ColumnName);
            }

            foreach (DataRow r in X.Rows)
            {
                DataRow A = Y.NewRow();
                for (int i = 0; i <= Y.Columns.Count - 1; i++)
                {
                    A[i] = r[i];
                }
                Y.Rows.Add(A);

                if (r["Type"].ToString() != "R")
                {
                    foreach (DataRow s in GetBPCSBOM(r["Item"].ToString(), EffectiveDate, Facility, level).Rows)
                    {
                        DataRow B = Y.NewRow();
                        for (int i = 0; i <= Y.Columns.Count - 1; i++)
                        {
                            B[i] = s[i];
                        }
                        Y.Rows.Add(B);
                    }
                }
            }

            if (Y.Rows.Count <= 0 && lvl == 0)
            {
                Y.Columns.Add("Error");
                Y.Rows.Add();
                Y.Rows[0][0] = "No records found.";
            }

            return Y;
        }
        #endregion

        #region SQL
        /// <summary>
        ///     Open the SQL connection with the Specified Database on SRVSQL01.
        /// </summary>
        /// <param name="DBName">
        ///     String Containing the Database name.
        /// </param>
        /// <returns>
        ///     Returns the Status of the Connection: 'OK' if the connection was successful or an Error Message if something went wrong.
        /// </returns>
        private string OpenSQLConn()
        {
            string bd = "";
            try
            {
                SQLConn.ConnectionString = " Data Source=SRVSQL01;Initial Catalog=" + DBName + ";User Id=" + DBUser + ";Password=" + DBPass;
                SQLConn.Open(); //'Opening connection
                bd = "OK";
            }
            catch (Exception e)
            {
                if (SQLConn.State != System.Data.ConnectionState.Closed)
                {
                    SQLConn.Close();
                }
                bd = "Error[Open_SQLConn]: " + e.Message;
            }
            return bd;
        }

        /// <summary>
        ///     Close the Connection of SQLConnection.
        /// </summary>
        private void CloseSQLConn()
        {
            if (SQLConn.State != System.Data.ConnectionState.Closed)
            {
                SQLConn.Close();
            }
        }

        /// <summary>
        ///     Executes an SQL command and return a DataTable with the result.
        /// </summary>
        /// <param name="SQLcmd">
        ///     Contains the SQLCommand with the parameters to execute.
        /// </param>
        /// <returns>
        ///     Returns a DataTable with the result or a failure message.
        /// </returns>
        public System.Data.DataTable GetSQLData(SqlCommand SQLcmd)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            SqlDataAdapter SQLda = new SqlDataAdapter();
            string st = OpenSQLConn();
            if (st == "OK")
            {
                try
                {
                    SQLcmd.Connection = SQLConn;
                    SQLda.SelectCommand = SQLcmd;
                    SQLda.Fill(dt);
                }
                catch (Exception e)
                {
                    dt = ErrorMsg("Error[GetSQLData]: " + e.Message);
                }
            }
            else
            {
                dt = ErrorMsg(st);
            }
            CloseSQLConn();
            return dt;
        }

        /// <summary>
        ///     Executes the command specified on SQLcmd.
        /// </summary>
        /// <param name="SQLcmd">
        ///     Contains the SQLCommand containing the parameters required to execute the SQL instruction.
        /// </param>
        /// <returns>
        ///     Returns a string containing the message either was a success or an error.
        /// </returns>
        public string ExecuteCommandSQL(SqlCommand SQLcmd)
        {
            string st = OpenSQLConn();
            if (st == "OK")
            {
                try
                {
                    SQLcmd.Connection = SQLConn;
                    SQLcmd.ExecuteNonQuery();
                    st = "The command has been executed without any problems.";
                }
                catch (Exception e)
                {
                    st = "Error[ExecuteCommandSQL]: " + e.Message + "\r\n\nPlease try to run the process again, if the problem persist, please contact your system administrator.";
                }
            }
            else
            {
                st = st + "\r\n\nPlease try to run the process again, if the problem persist, please contact your system administrator.";
            }
            CloseSQLConn();
            return st;
        }

        public string DBName
        {
            get { return dbname; }
            set
            {
                dbname = value;
            }
        }

        public string DBUser
        {
            get { return dbuser; }
            set
            {
                dbuser = value;
            }
        }

        public string DBPass
        {
            get { return dbpass; }
            set
            {
                dbpass = value;
            }
        }
        #endregion

        #region Error Message
        /// <summary>
        ///     Gets the error message and transforms it into a DataTable to be handled properly.
        /// </summary>
        /// <param name="fun">
        ///     Contains the error message to display.
        /// </param>
        /// <returns>
        ///     Returns a DataTable with the error message.
        /// </returns>
        public static System.Data.DataTable ErrorMsg(string fun)
        {
            System.Data.DataTable ret = new System.Data.DataTable();
            System.Data.DataRow dr = ret.NewRow();
            dr[1] = fun + "\r\n\n" + "Please try to run the process again, if the problem persist, please contact your system administrator.";
            ret.Columns.Add("Error");
            ret.Rows.Add(dr);
            return ret;
        }
        #endregion
    }
    #endregion

    #region DataGridPlus

    #region Data gridPlus Class
    public class DatagridPlus : DataGridView
    {
        private bool rhs = false;

        private AdvancedProperties X;

        /// <summary>
        /// Constructor Class
        /// </summary>
        [System.ComponentModel.Browsable(true), System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Always)]
        public DatagridPlus()
        {
            #region Handlers
            this.DataBindingComplete += new System.Windows.Forms.DataGridViewBindingCompleteEventHandler(this.G_DataBindingComplete);
            this.RowHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.G_RowHeaderMouseClick);
            #endregion

            X = new AdvancedProperties();
            this.AllowUserToAddRows = false;
        }

        #region Properties
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        [TypeConverter(typeof(AdvancedPropertiesConverter))]
        [System.ComponentModel.Description("Advanced properties used in the DataGridPlus.")]
        public AdvancedProperties AdvancedProperties
        {
            get { return X; }
            set
            {
                X = value;
            }
        }
        #endregion

        #region Events
        private void G_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            bool s = true;
            switch (this.X.GroupingMode)
            {
                case AdvancedProperties.GroupMethod.None:
                    if (AdvancedProperties.EnableAlternateColor)
                    {
                        foreach (DataGridViewRow row in base.Rows)
                        {
                            if (s)
                            {
                                row.DefaultCellStyle.BackColor = AdvancedProperties.AlternateRowColor1;
                                s = false;
                            }
                            else
                            {
                                row.DefaultCellStyle.BackColor = AdvancedProperties.AlternateRowColor2;
                                s = true;
                            }
                        }
                    }
                    break;
                case AdvancedProperties.GroupMethod.Custom:
                    string current = "";

                    this.Columns[0].Visible = X.ShowFirstColumns;

                    this.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

                    for (int i = 0; i < this.Rows.Count; i++)
                    {
                        if (current != this.Rows[i].Cells[0].Value.ToString())
                        {
                            current = this.Rows[i].Cells[0].Value.ToString();
                            this.Rows[i].HeaderCell.Value = "[+]";
                            this.Rows[i].DefaultCellStyle.BackColor = X.ParentColor;

                            if (s)
                            {
                                this.Rows[i].DefaultCellStyle.BackColor = AdvancedProperties.AlternateRowColor1;
                                s = false;
                            }
                            else
                            {
                                this.Rows[i].DefaultCellStyle.BackColor = AdvancedProperties.AlternateRowColor2;
                                s = true;
                            }

                        }
                        else
                        {
                            this.Rows[i].Visible = false;
                            this.Rows[i].HeaderCell.Value = "L";
                            this.Rows[i].DefaultCellStyle.BackColor = X.ChildColor;
                        }
                    }

                    break;
                case AdvancedProperties.GroupMethod.Constant:
                    if (this.X.AgroupationNumber > 0)
                    {
                        int a = 0;
                        foreach (DataGridViewRow row in this.Rows)
                        {
                            if (row.Index == a)
                            {
                                if (row.Index != this.Rows.Count - 1)
                                {
                                    row.HeaderCell.Value = "[+]";
                                    row.DefaultCellStyle.BackColor = X.ParentColor;

                                    if (s)
                                    {
                                        row.DefaultCellStyle.BackColor = AdvancedProperties.AlternateRowColor1;
                                        s = false;
                                    }
                                    else
                                    {
                                        row.DefaultCellStyle.BackColor = AdvancedProperties.AlternateRowColor2;
                                        s = true;
                                    }

                                    a = a + this.X.AgroupationNumber;
                                }
                            }
                            else
                            {
                                row.Visible = false;
                                row.HeaderCell.Value = "L";
                                row.DefaultCellStyle.BackColor = X.ChildColor;
                            }
                        }
                    }
                    break;
            }

            this.RowHeadersWidth = 60;
        }

        private void G_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            switch ((string)this.Rows[e.RowIndex].HeaderCell.Value)
            {
                case "[+]":
                    rhs = false;
                    this.Rows[e.RowIndex].HeaderCell.Value = "[-]";
                    break;
                case "[-]":
                    rhs = true;
                    this.Rows[e.RowIndex].HeaderCell.Value = "[+]";
                    break;
            }

            switch (this.X.GroupingMode)
            {
                case AdvancedProperties.GroupMethod.None:
                    //Do Nothing
                    break;
                case AdvancedProperties.GroupMethod.Custom:
                    string current = "";
                    bool first = true;

                    for (int i = e.RowIndex; i <= this.Rows.Count - 1; i++)
                    {
                        if (current != this.Rows[i].Cells[0].Value.ToString())
                        {
                            if (first)
                            {
                                current = this.Rows[e.RowIndex].Cells[0].Value.ToString();
                                first = false;
                            }
                            else
                            {
                                break;
                            }
                        }
                        else
                        {
                            if (rhs == false)
                            {
                                this.Rows[i].Visible = true;
                            }
                            else
                            {
                                this.Rows[i].Visible = false;
                            }
                        }
                    }
                    break;
                case AdvancedProperties.GroupMethod.Constant:
                    for (int i = e.RowIndex + 1; i < e.RowIndex + this.X.AgroupationNumber; i++)
                    {
                        if (i > this.Rows.Count - 1)
                        {
                            break;
                        }
                        else
                        {
                            if (rhs == false)
                            {
                                this.Rows[i].Visible = true;
                            }
                            else if (rhs == true)
                            {
                                this.Rows[i].Visible = false;
                            }
                        }
                    }
                    break;
            }
        }
        #endregion
    }
    #endregion

    #region AdvancedProperties Class
    public class AdvancedProperties
    {
        #region Variable Declaration
        private GroupMethod GM = GroupMethod.None;
        private int AN = 0;
        private Color Pcolor = Color.White;
        private Color Ccolor = Color.White;
        private Color Alt1 = Color.White;
        private Color Alt2 = Color.SlateGray;
        bool altcolor = false;
        private bool SFC = false;
        #endregion

        #region Constructor
        public AdvancedProperties()
        {

        }
        #endregion

        #region Porperties
        [System.ComponentModel.Browsable(true), System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Always)]
        [System.ComponentModel.Description("How the rows will be grouped.")]
        public GroupMethod GroupingMode
        {
            get { return GM; }
            set
            {
                if (value != GM)
                {
                    GM = value;
                }
            }
        }

        [System.ComponentModel.Description("Number of rows that will be grouped ONLY when GroupingMode is set to 'Constant'.")]
        [System.ComponentModel.Browsable(true), System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Always)]
        public int AgroupationNumber
        {
            get { return AN; }
            set
            {
                if (value != AN)
                {
                    AN = value;
                }
            }
        }

        [System.ComponentModel.Description("BackColor of the parent Row.")]
        [System.ComponentModel.Browsable(true), System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Always)]
        public Color ParentColor
        {
            get { return Pcolor; }
            set
            {
                if (value != Pcolor)
                {
                    Pcolor = value;
                }
            }
        }

        [System.ComponentModel.Description("BackColor of the Child Rows.")]
        [System.ComponentModel.Browsable(true), System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Always)]
        public Color ChildColor
        {
            get { return Ccolor; }
            set
            {
                if (value != Ccolor)
                {
                    Ccolor = value;
                }
            }
        }

        [System.ComponentModel.Description("Shows the first column if set to 'TRUE' or is hidden when is set to 'FALSE'. Useful when using the 'Custom' GroupingMode.")]
        [System.ComponentModel.Browsable(true), System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Always)]
        public bool ShowFirstColumns
        {
            get { return SFC; }
            set
            {
                if (value != SFC)
                {
                    SFC = value;
                }
            }
        }

        [System.ComponentModel.Description("Color for the Odd-Rows when EnableAlternateColor is true.")]
        [System.ComponentModel.Browsable(true), System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Always)]
        public Color AlternateRowColor1
        {
            get { return Alt1; }
            set
            {
                Alt1 = value;
            }
        }

        [System.ComponentModel.Description("Color for the Even-Rows when EnableAlternateColor is true.")]
        [System.ComponentModel.Browsable(true), System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Always)]
        public Color AlternateRowColor2
        {
            get { return Alt2; }
            set
            {
                Alt2 = value;
            }
        }

        [System.ComponentModel.Description("Enables Alternate row colors on the rows. If is set to true, it will override the ParentColor Property.")]
        [System.ComponentModel.Browsable(true), System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Always)]
        public bool EnableAlternateColor
        {
            get { return altcolor; }
            set
            {
                if (value != altcolor)
                {
                    altcolor = value;
                }
            }
        }
        #endregion

        #region Enumerations
        public enum GroupMethod
        {
            None = 0,
            Custom = 1,
            Constant = 2
        }
        #endregion

    }

    #region Type Converter Class
    public class AdvancedPropertiesConverter : ExpandableObjectConverter
    {
        public override object ConvertTo(
                 ITypeDescriptorContext context,
                 CultureInfo culture,
                 object value,
                 Type destinationType)
        {
            if (destinationType == typeof(string))
            {
                return ((AdvancedProperties)value).AgroupationNumber.ToString() + "," +
                       ((AdvancedProperties)value).AlternateRowColor1.Name.ToString() + "," +
                       ((AdvancedProperties)value).AlternateRowColor2.Name.ToString() + "," +
                       ((AdvancedProperties)value).ChildColor.Name.ToString() + "," +
                       ((AdvancedProperties)value).EnableAlternateColor.ToString() + "," +
                       ((AdvancedProperties)value).GroupingMode.ToString() + "," +
                       ((AdvancedProperties)value).ParentColor.Name.ToString() + "," +
                       ((AdvancedProperties)value).ShowFirstColumns.ToString();
            }

            return base.ConvertTo(
                context,
                culture,
                value,
                destinationType);
        }
    }
    #endregion

    #endregion

    #endregion

    #region Centered MessageBox
    public class Centered_MessageBox : IDisposable
    {
        private int mTries = 0;
        private Form mOwner;

        public Centered_MessageBox(Form owner)
        {
            mOwner = owner;
            owner.BeginInvoke(new MethodInvoker(findDialog));
        }

        private void findDialog()
        {
            if (mTries < 0) return;
            EnumThreadWndProc callback = new EnumThreadWndProc(checkWindow);
            if (EnumThreadWindows(GetCurrentThreadId(), callback, IntPtr.Zero))
            {
                if (++mTries < 10) mOwner.BeginInvoke(new MethodInvoker(findDialog));
            }
        }
        private bool checkWindow(IntPtr hWnd, IntPtr lp)
        {
            StringBuilder sb = new StringBuilder(260);
            GetClassName(hWnd, sb, sb.Capacity);
            if (sb.ToString() != "#32770") return true;

            Rectangle frmRect = new Rectangle(mOwner.Location, mOwner.Size);
            RECT dlgRect;
            GetWindowRect(hWnd, out dlgRect);
            MoveWindow(hWnd,
                frmRect.Left + (frmRect.Width - dlgRect.Right + dlgRect.Left) / 2,
                frmRect.Top + (frmRect.Height - dlgRect.Bottom + dlgRect.Top) / 2,
                dlgRect.Right - dlgRect.Left,
                dlgRect.Bottom - dlgRect.Top, true);
            return false;
        }
        public void Dispose()
        {
            mTries = -1;
        }

        // P/Invoke declarations
        private delegate bool EnumThreadWndProc(IntPtr hWnd, IntPtr lp);
        [DllImport("user32.dll")]
        private static extern bool EnumThreadWindows(int tid, EnumThreadWndProc callback, IntPtr lp);
        [DllImport("kernel32.dll")]
        private static extern int GetCurrentThreadId();
        [DllImport("user32.dll")]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder buffer, int buflen);
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT rc);
        [DllImport("user32.dll")]
        private static extern bool MoveWindow(IntPtr hWnd, int x, int y, int w, int h, bool repaint);
        private struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    }
    #endregion

    #region GrayTabControl
    class GrayTabControl : TabControlDesigner
    {
        private Color C1, C2, C3, C4, C5, C6;
        private LinearGradientBrush L1, L2;

        public GrayTabControl()
        {
            C1 = RGB(191, 191, 191); //Panel Color
            C2 = RGB(169, 169, 169); //Line Color
            C3 = ARGB(150, Color.White); //Highlight
            C4 = RGB(169, 169, 169); //Unselected Tab #1
            C5 = RGB(169, 169, 169); //Unselected Tab #2
            C6 = RGB(169, 169, 169); //Tab Line

            Alignment = TabAlignment.Top;
            Font = new Font("Verdana", 8, FontStyle.Bold);
            ItemSize = new Size(120, 30);
            PanelColor = Color.FromArgb(169, 169, 169);

            for (int i = 0; i <= TabPages.Count - 1; i++)
            {
                TabPages[i].BackColor = C1;
            }
        }

        protected override void TabPaint(int e)
        {
            string Text = TabPages[e].Text;
            Rectangle Tab = GetTabRect(e);
            Rectangle Temp, Highlight;
            Point[] Outline;
            int ind;


            if (SelectedIndex == e)
            {
                ind = 1;
            }
            else
            {
                ind = 0;
            }

            Temp = new Rectangle(Tab.X + 1, Tab.Y + 1, Tab.Width - 6, Tab.Height + ind);
            Highlight = new Rectangle(Temp.X, Temp.Y, Temp.Width, 10);
            Outline = new Point[] { new Point(Temp.X, Temp.Bottom), new Point(Temp.X, Temp.Top), new Point(Temp.Right, Temp.Top), new Point(Temp.Right, Temp.Bottom) };

            L1 = new LinearGradientBrush(Highlight, Color.FromArgb(128, 147, 255), Color.Transparent, 90);
            L2 = new LinearGradientBrush(Temp, C4, C5, 90);

            if (SelectedIndex == e)
            {
                try
                {
                    TabPages[e].BackColor = C1;
                }
                catch { }
                G.FillRectangle(ToBrush(C1), Temp);

            }
            else if (State == MouseState.Over && Temp.Contains(Coordinates))
            {
                G.FillRectangle(L2, Temp);
                G.FillRectangle(L1, Highlight);
            }
            else
            {
                G.FillRectangle(L2, Temp);
            }

            G.DrawLines(ToPen(C2), Outline);
            G.DrawString(Text, Font, Brushes.Blue, Center(Text, Font, Temp, 0, 0));
        }

        protected override void PaintHook()
        {
            Point[] Outline = new Point[]{new Point(ClientRectangle.X + 3, ClientRectangle.Y + 34),
                                          new Point(ClientRectangle.X + 3, ClientRectangle.Bottom - 4),
                                          new Point(ClientRectangle.Right - 4, ClientRectangle.Bottom - 4),
                                          new Point(ClientRectangle.Right - 4, ClientRectangle.Top + 34)};
            G.DrawLines(ToPen(C6), Outline);
            G.DrawLine(ToPen(C6), 4, 33, Width - 4, 33);
        }
    }
    #endregion

    #region VerticalTabControl
    class VerticalTabControl : TabControlDesigner
    {
        private Color C1, C2, C3, C4, C5, C6;

        public VerticalTabControl()
        {
            C1 = Color.FromArgb(103, 103, 103);
            C2 = Color.FromArgb(81, 81, 81);
            C3 = Color.FromArgb(50, 50, 50);
            C4 = Color.FromArgb(120, 120, 120);
            C5 = Color.FromArgb(150, 150, 150);
            C6 = Color.FromArgb(50, Color.Black);

            Alignment = TabAlignment.Left;
            ItemSize = new Size(44, 136);
            SizeMode = TabSizeMode.Fixed;
            ItemSize = new Size(40, 110);
            Font = new Font("Century Gothic", 11, FontStyle.Bold);
            PanelColor = C1;

            for (int i = 0; i <= TabPages.Count - 1; i++)
            {
                TabPages[i].BackColor = Color.DarkGray;
            }
        }

        protected override void TabPaint(int e)
        {
            string T = TabPages[e].Text;
            Rectangle R = GetTabRect(e);
            Rectangle Temp = new Rectangle(R.X + 2, R.Y + 5, 110, 35);

            if (SelectedIndex == e)
            {
                try
                {
                    TabPages[e].BackColor = Color.DarkGray;
                }
                catch { }

                G.FillRectangle(ToBrush(C2), Temp);
                G.DrawRectangle(ToPen(C3), Temp);

                int X = Temp.Right;
                int Y = Temp.Y + Temp.Height / 2;

                Point[] P = new Point[] { new Point(X, Y - 5), new Point(X, Y + 5), new Point(X - 5, Y) };

                G.FillPolygon(Brushes.WhiteSmoke, P);

            }
            else if (State == MouseState.Over && Temp.Contains(Coordinates))
            {
                G.FillRectangle(ToBrush(C4), Temp);
                G.DrawRectangle(ToPen(C5), Temp);
            }

            G.DrawString(T, Font, Brushes.Black, Center(T, Font, Temp, 0, -1));
            G.DrawString(T, Font, Brushes.White, Center(T, Font, Temp));
        }
    }
    #endregion

    #region CustomTab

    //'-------------------------------------
    //'TabControlDesigner
    //'
    //This is the translated code to C# by Z3r0Cr45h
    //
    //'Creator(VB.NET): Eprouvez
    //'Date: 6/5/2012
    //'Updated: 6/6/2012
    //'Version: 1.1.1
    //'
    //'Credits to:
    //'Aeonhack
    //'mavamaarten
    //'--------------------------------------

    #region Bloom
    public class Bloom
    {
        public Bloom(string Name, Color color)
        {
            _Name = Name;
            _Value = color;
        }
        public Bloom(string Name, int Red, int Green, int Blue)
        {
            _Name = Name;
            _Value = Color.FromArgb(Red, Green, Blue);
        }
        public Bloom(string Name, int Alpha, int Red, int Green, int Blue)
        {
            _Name = Name;
            _Value = Color.FromArgb(Alpha, Red, Green, Blue);
        }

        private readonly string _Name;
        public string Name
        {
            get { return _Name; }
        }

        private Color _Value;
        public Color Color
        {
            get { return _Value; }
            set
            {
                _Value = value;
            }
        }
    }
    #endregion

    #region Tab control Designer Class
    abstract class TabControlDesigner : TabControl
    {
        protected Graphics G;
        protected MouseState State;
        protected Point Coordinates;

        #region Enumeration
        protected enum MouseState
        {
            None = 0,
            Over = 1,
            Down = 2
        }
        #endregion

        #region Routines

        public TabControlDesigner()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.ResizeRedraw | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.SupportsTransparentBackColor, true);
            DoubleBuffered = true;
            SizeMode = TabSizeMode.Fixed;
            ItemSize = new Size(250, 26);
        }

        protected void SetState(MouseState current)
        {
            State = current;
            Invalidate();
        }
        #endregion

        #region Hooks
        protected abstract void TabPaint(int Index);
        protected virtual void ColorHook() { }
        protected virtual void PaintHook() { }
        #endregion

        #region Properties

        private int _InactiveIconOpacity = 50;
        public int InactiveIconOpacity
        {
            get { return _InactiveIconOpacity; }
            set
            {
                if (value == 0) { value = 50; }
                if (value < 0) { value = 0; }
                if (value > 100) { value = 100; }
                _InactiveIconOpacity = value;
            }
        }

        private Color _PanelColor;
        public Color PanelColor
        {
            get { return _PanelColor; }
            set
            {
                _PanelColor = value;
            }
        }

        private Dictionary<string, Color> Items = new Dictionary<string, Color>();
        public Bloom[] Colors
        {
            get
            {
                List<Bloom> T = new List<Bloom>();
                Dictionary<string, Color>.Enumerator E = Items.GetEnumerator();
                while (E.MoveNext())
                {
                    T.Add(new Bloom(E.Current.Key, E.Current.Value));
                }
                return T.ToArray();
            }
            set
            {
                foreach (Bloom B in value)
                {
                    if (Items.ContainsKey(B.Name)) { Items[B.Name] = B.Color; }
                }

                InvalidateCustomization();
                ColorHook();
                Invalidate();
            }
        }

        private string _Customization;
        public string Customization
        {
            get { return _Customization; }
            set
            {
                if (value == _Customization)
                {
                    return;
                }

                byte[] Data;
                Bloom[] Items = Colors;

                try
                {
                    Data = Convert.FromBase64String(value);
                    for (int I = 0; I <= Items.Length - 1; I++)
                    {
                        Items[I].Color = Color.FromArgb(BitConverter.ToInt32(Data, I * 4));
                    }
                }
                catch
                {
                    return;
                }

                _Customization = value;
                Colors = Items;
                ColorHook();
                Invalidate();
            }
        }
        #endregion

        #region PropertyHelpers
        protected Pen GetPen(string name)
        {
            return new Pen(Items[name]);
        }
        protected Pen GetPen(string name, Single width)
        {
            return new Pen(Items[name], width);
        }
        protected SolidBrush GetBrush(string name)
        {
            return new SolidBrush(Items[name]);
        }
        protected Color GetColor(string name)
        {
            return Items[name];
        }
        protected void SetColor(string name, Color value)
        {
            if (Items.ContainsKey(name)) { Items[name] = value; } else { Items.Add(name, value); }
        }
        protected void SetColor(string name, byte r, byte g, byte b)
        {
            SetColor(name, Color.FromArgb(r, g, b));
        }
        protected void SetColor(string name, byte a, byte r, byte g, byte b)
        {
            SetColor(name, Color.FromArgb(a, r, g, b));
        }
        protected void SetColor(string name, byte a, Color value)
        {
            SetColor(name, Color.FromArgb(a, value));
        }
        private void InvalidateCustomization()
        {
            MemoryStream M = new MemoryStream(Items.Count * 4);

            foreach (Bloom B in Colors)
            {
                M.Write(BitConverter.GetBytes(B.Color.ToArgb()), 0, 4);
            }

            M.Close();
            _Customization = Convert.ToBase64String(M.ToArray());
        }
        #endregion

        #region DrawingMethods
        public SizeF Measure(string Text)
        {
            return G.MeasureString(Text, Font);
        }
        public SizeF Measure(string Text, Font Font)
        {
            return G.MeasureString(Text, Font);
        }
        public Point Center(string Text, Rectangle Area)
        {
            return Center(Text, Font, Area);
        }
        public Point Center(string Text, Font Font, Rectangle Area)
        {
            return Center(Text, Font, Area, 0, 0);
        }
        public Point Center(string Text, Font Font, Rectangle Area, int XOffset, int YOffset)
        {
            SizeF M = Measure(Text, Font);
            return new Point(Convert.ToInt32(Area.X + Area.Width / 2 - M.Width / 2) + XOffset, Convert.ToInt32(Area.Y + Area.Height / 2 - M.Height / 2 + YOffset));
        }
        public Pen ToPen(Color color)
        {
            return new Pen(color);
        }
        public Brush ToBrush(Color color)
        {
            return new SolidBrush(color);
        }
        public Color RGB(int Red, int Green, int Blue)
        {
            return Color.FromArgb(Red, Green, Blue);
        }
        public Color ARGB(int Alpha, Color color)
        {
            return Color.FromArgb(Alpha, color);
        }
        public Color ARGB(int Alpha, int Red, int Green, int Blue)
        {
            return Color.FromArgb(Alpha, Red, Green, Blue);
        }
        public Rectangle Shrink(Rectangle rectangle, int Offset)
        {
            return Shrink(rectangle, Offset, true);
        }
        public Rectangle Shrink(Rectangle rectangle, int Offset, bool CenterPoint)
        {
            int O;
            if (CenterPoint) { O = Offset; } else { O = 0; }
            Rectangle R = new Rectangle(rectangle.X + O, rectangle.Y + O, rectangle.Width - Offset * 2, rectangle.Height - Offset * 2);
            return R;
        }
        public Rectangle Enlarge(Rectangle rectangle, int Offset)
        {
            return Enlarge(rectangle, Offset, true);
        }
        public Rectangle Enlarge(Rectangle rectangle, int Offset, bool CenterPoint)
        {
            int O;
            if (CenterPoint) { O = Offset; } else { O = 0; }
            Rectangle R = new Rectangle(rectangle.X + O, rectangle.Y + O, rectangle.Width - Offset * 2, rectangle.Height + Offset * 2);
            return R;
        }
        public Image ImageOpacity(Bitmap Image, Single Opacity)
        {
            Bitmap Result = new Bitmap(Image.Width, Image.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);

            Opacity = Math.Min(Opacity, 100);
            System.Drawing.Imaging.ImageAttributes Attributes = new System.Drawing.Imaging.ImageAttributes();
            using (Attributes)
            {
                ColorMatrix Matrix = new ColorMatrix();
                PointF[] Points = { new Point(0, 0), new Point(Image.Width, 0), new Point(0, Image.Height) };
                Graphics i = Graphics.FromImage(Result);

                Matrix.Matrix33 = Opacity / 100.0F;
                Attributes.SetColorMatrix(Matrix);

                using (i)
                {
                    i.Clear(Color.Transparent);
                    i.DrawImage(Image, Points, new RectangleF(Point.Empty, Image.Size), GraphicsUnit.Pixel, Attributes);
                }
            }

            return Result;
        }
        #endregion

        #region OverrideMethods
        protected override void CreateHandle()
        {
            base.CreateHandle();
            base.DoubleBuffered = true;
            InvalidateCustomization();
            ColorHook();
            SizeMode = TabSizeMode.Fixed;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.ResizeRedraw | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.SupportsTransparentBackColor, true);
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            ColorHook();
            base.OnHandleCreated(e);
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            SetState(MouseState.Over);
            base.OnMouseEnter(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            SetState(MouseState.None);
            for (int i = 0; i <= TabPages.Count - 1; i++)
            {
                if (TabPages[i].DisplayRectangle.Contains(Coordinates))
                {
                    base.Invalidate();
                    break;
                }
            }
            base.OnMouseLeave(e);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            Coordinates = e.Location;
            for (int i = 0; i <= TabPages.Count - 1; i++)
            {
                if (TabPages[i].DisplayRectangle.Contains(Coordinates))
                {
                    base.Invalidate();
                    break;
                }
            }
            base.OnMouseMove(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            G = e.Graphics;

            if (null != _PanelColor)
            {
                G.FillRectangle(ToBrush(_PanelColor), e.ClipRectangle);
            }

            PaintHook();

            for (int i = 0; i <= TabPages.Count - 1; i++)
            {
                TabPaint(i);
            }

            base.OnPaint(e);
        }
        #endregion

    }
    #endregion
    #endregion

    #region ExcelEnumerations
    public class Excel
    {
        #region xlConstants
        public enum xlConstants
        {
            xl3DBar = -4099,
            xl3DEffects1 = 13,
            xl3DEffects2 = 14,
            xl3DSurface = -4103,
            xlAbove = 0,
            xlAccounting1 = 4,
            xlAccounting2 = 5,
            xlAccounting4 = 17,
            xlAdd = 2,
            xlAll = -4104,
            xlAccounting3 = 6,
            xlAllExceptBorders = 7,
            xlAutomatic = -4105,
            xlBar = 2,
            xlBelow = 1,
            xlBidi = -5000,
            xlBidiCalendar = 3,
            xlBoth = 1,
            xlBottom = -4107,
            xlCascade = 7,
            xlCenter = -4108,
            xlCenterAcrossSelection = 7,
            xlChart4 = 2,
            xlChartSeries = 17,
            xlChartShort = 6,
            xlChartTitles = 18,
            xlChecker = 9,
            xlCircle = 8,
            xlClassic1 = 1,
            xlClassic2 = 2,
            xlClassic3 = 3,
            xlClosed = 3,
            xlColor1 = 7,
            xlColor2 = 8,
            xlColor3 = 9,
            xlColumn = 3,
            xlCombination = -4111,
            xlComplete = 4,
            xlConstants = 2,
            xlContents = 2,
            xlContext = -5002,
            xlCorner = 2,
            xlCrissCross = 16,
            xlCross = 4,
            xlCustom = -4114,
            xlDebugCodePane = 13,
            xlDefaultAutoFormat = -1,
            xlDesktop = 9,
            xlDiamond = 2,
            xlDirect = 1,
            xlDistributed = -4117,
            xlDivide = 5,
            xlDoubleAccounting = 5,
            xlDoubleClosed = 5,
            xlDoubleOpen = 4,
            xlDoubleQuote = 1,
            xlDrawingObject = 14,
            xlEntireChart = 20,
            xlExcelMenus = 1,
            xlExtended = 3,
            xlFill = 5,
            xlFirst = 0,
            xlFixedValue = 1,
            xlFloating = 5,
            xlFormats = -4122,
            xlFormula = 5,
            xlFullScript = 1,
            xlGeneral = 1,
            xlGray16 = 17,
            xlGray25 = -4124,
            xlGray50 = -4125,
            xlGray75 = -4126,
            xlGray8 = 18,
            xlGregorian = 2,
            xlGrid = 15,
            xlGridline = 22,
            xlHigh = -4127,
            xlHindiNumerals = 3,
            xlIcons = 1,
            xlImmediatePane = 12,
            xlInside = 2,
            xlInteger = 2,
            xlJustify = -4130,
            xlLast = 1,
            xlLastCell = 11,
            xlLatin = -5001,
            xlLeft = -4131,
            xlLeftToRight = 2,
            xlLightDown = 13,
            xlLightHorizontal = 11,
            xlLightUp = 14,
            xlLightVertical = 12,
            xlList1 = 10,
            xlList2 = 11,
            xlList3 = 12,
            xlLocalFormat1 = 15,
            xlLocalFormat2 = 16,
            xlLogicalCursor = 1,
            xlLong = 3,
            xlLotusHelp = 2,
            xlLow = -4134,
            xlLTR = -5003,
            xlMacrosheetCell = 7,
            xlManual = -4135,
            xlMaximum = 2,
            xlMinimum = 4,
            xlMinusValues = 3,
            xlMixed = 2,
            xlMixedAuthorizedScript = 4,
            xlMixedScript = 3,
            xlModule = -4141,
            xlMultiply = 4,
            xlNarrow = 1,
            xlNextToAxis = 4,
            xlNoDocuments = 3,
            xlNone = -4142,
            xlNotes = -4144,
            xlOff = -4146,
            xlOn = 1,
            xlOpaque = 3,
            xlOpen = 2,
            xlOutside = 3,
            xlPartial = 3,
            xlPartialScript = 2,
            xlPercent = 2,
            xlPlus = 9,
            xlPlusValues = 2,
            xlReference = 4,
            xlRight = -4152,
            xlRTL = -5004,
            xlScale = 3,
            xlSemiautomatic = 2,
            xlSemiGray75 = 10,
            xlShort = 1,
            xlShowLabel = 4,
            xlShowLabelAndPercent = 5,
            xlShowPercent = 3,
            xlShowValue = 2,
            xlSimple = -4154,
            xlSingle = 2,
            xlSingleAccounting = 4,
            xlSingleQuote = 2,
            xlSquare = 1,
            xlStar = 5,
            xlStError = 4,
            xlStrict = 2,
            xlSubtract = 3,
            xlSystem = 1,
            xlTextBox = 16,
            xlTiled = 1,
            xlTitleBar = 8,
            xlToolbar = 1,
            xlToolbarButton = 2,
            xlTop = -4160,
            xlTopToBottom = 1,
            xlTransparent = 2,
            xlTriangle = 3,
            xlVeryHidden = 2,
            xlVisible = 12,
            xlVisualCursor = 2,
            xlWatchPane = 11,
            xlWide = 3,
            xlWorkbookTab = 6,
            xlWorksheet4 = 1,
            xlWorksheetCell = 3,
            xlWorksheetShort = 5,
        }
        #endregion

        #region xlAboveBelow
        public enum xlAboveBelow
        {
            XlAboveAverage = 0,
            XlAboveStdDev = 4,
            XlBelowAverage = 1,
            XlBelowStdDev = 5,
            XlEqualAboveAverage = 2,
            XlEqualBelowAverage = 3,
        }
        #endregion

        #region xlActionType
        public enum xlActionType
        {
            xlActionTypeDrillthrough = 256,
            xlActionTypeReport = 128,
            xlActionTypeRowset = 16,
            xlActionTypeUrl = 1,
        }
        #endregion

        #region xlAllocation
        public enum xlAllocation
        {
            xlAutomaticAllocation = 2,
            xlManualAllocation = 1,
        }
        #endregion

        #region xlAllocationMethod
        public enum xlAllocationMethod
        {
            xlEqualAllocation = 1,
            xlWeightedAllocation = 2,
        }
        #endregion

        #region xlAllocationValue
        public enum xlAllocationValue
        {
            xlAllocateIncrement = 2,
            xlAllocateValue = 1,
        }
        #endregion

        #region XlApplicationInternational
        public enum XlApplicationInternational
        {
            xl24HourClock = 33,
            xl4DigitYears = 43,
            xlAlternateArraySeparator = 16,
            xlColumnSeparator = 14,
            xlCountryCode = 1,
            xlCountrySetting = 2,
            xlCurrencyBefore = 37,
            xlCurrencyCode = 25,
            xlCurrencyDigits = 27,
            xlCurrencyLeadingZeros = 40,
            xlCurrencyMinusSign = 38,
            xlCurrencyNegative = 28,
            xlCurrencySpaceBefore = 36,
            xlCurrencyTrailingZeros = 39,
            xlDateOrder = 32,
            xlDateSeparator = 17,
            xlDayCode = 21,
            xlDayLeadingZero = 42,
            xlDecimalSeparator = 3,
            xlGeneralFormatName = 26,
            xlHourCode = 22,
            xlLeftBrace = 12,
            xlLeftBracket = 10,
            xlListSeparator = 5,
            xlLowerCaseColumnLetter = 9,
            xlLowerCaseRowLetter = 8,
            xlMDY = 44,
            xlMetric = 35,
            xlMinuteCode = 23,
            xlMonthCode = 20,
            xlMonthLeadingZero = 41,
            xlMonthNameChars = 30,
            xlNoncurrencyDigits = 29,
            xlNonEnglishFunctions = 34,
            xlRightBrace = 13,
            xlRightBracket = 11,
            xlRowSeparator = 15,
            xlSecondCode = 24,
            xlThousandsSeparator = 4,
            xlTimeLeadingZero = 45,
            xlTimeSeparator = 18,
            xlUpperCaseColumnLetter = 7,
            xlUpperCaseRowLetter = 6,
            xlWeekdayNameChars = 31,
            xlYearCode = 19,
        }
        #endregion

        #region XlApplyNamesOrder
        public enum XlApplyNamesOrder
        {
            xlColumnThenRow = 2,
            xlRowThenColumn = 1,
        }
        #endregion

        #region XlArabicModes
        public enum XlArabicModes
        {
            xlArabicBothStrict = 3,
            xlArabicNone = 0,
            xlArabicStrictAlefHamza = 1,
            xlArabicStrictFinalYaa = 2,
        }
        #endregion

        #region XlArrangeStyle
        public enum XlArrangeStyle
        {
            xlArrangeStyleCascade = 7,
            xlArrangeStyleHorizontal = -4128,
            xlArrangeStyleTiled = 1,
            xlArrangeStyleVertical = -4166,
        }
        #endregion

        #region XlArrowHeadLength
        public enum XlArrowHeadLength
        {
            xlArrowHeadLengthLong = 3,
            xlArrowHeadLengthMedium = -4138,
            xlArrowHeadLengthShort = 1,
        }
        #endregion

        #region XlArrowHeadStyle
        public enum XlArrowHeadStyle
        {
            xlArrowHeadStyleClosed = 3,
            xlArrowHeadStyleDoubleClosed = 5,
            xlArrowHeadStyleDoubleOpen = 4,
            xlArrowHeadStyleNone = -4142,
            xlArrowHeadStyleOpen = 2,
        }
        #endregion

        #region XlArrowHeadWidth
        public enum XlArrowHeadWidth
        {
            xlArrowHeadWidthMedium = -4138,
            xlArrowHeadWidthNarrow = 1,
            xlArrowHeadWidthWide = 3,
        }
        #endregion

        #region XlAutoFillType
        public enum XlAutoFillType
        {
            xlFillCopy = 1,
            xlFillDays = 5,
            xlFillDefault = 0,
            xlFillFormats = 3,
            xlFillMonths = 7,
            xlFillSeries = 2,
            xlFillValues = 4,
            xlFillWeekdays = 6,
            xlFillYears = 8,
            xlGrowthTrend = 10,
            xlLinearTrend = 9,
        }
        #endregion

        #region XlAutoFilterOperator
        public enum XlAutoFilterOperator
        {
            xlAnd = 1,
            xlBottom10Items = 4,
            xlBottom10Percent = 6,
            xlOr = 2,
            xlTop10Items = 3,
            xlTop10Percent = 5,
        }
        #endregion

        #region XlAxisCrosses
        public enum XlAxisCrosses
        {
            xlAxisCrossesAutomatic = -4105,
            xlAxisCrossesCustom = -4114,
            xlAxisCrossesMaximum = 2,
            xlAxisCrossesMinimum = 4,
        }
        #endregion

        #region XlAxisGroup
        public enum XlAxisGroup
        {
            xlPrimary = 1,
            xlSecondary = 2,
        }
        #endregion

        #region XlAxisType
        public enum XlAxisType
        {
            xlCategory = 1,
            xlSeriesAxis = 3,
            xlValue = 2,
        }
        #endregion

        #region XlBackground
        public enum XlBackground
        {
            xlBackgroundAutomatic = -4105,
            xlBackgroundOpaque = 3,
            xlBackgroundTransparent = 2,
        }
        #endregion

        #region XlBarShape
        public enum XlBarShape
        {
            xlBox = 0,
            xlConeToMax = 5,
            xlConeToPoint = 4,
            xlCylinder = 3,
            xlPyramidToMax = 2,
            xlPyramidToPoint = 1,
        }
        #endregion

        #region XlBordersIndex
        public enum XlBordersIndex
        {
            xlDiagonalDown = 5,
            xlDiagonalUp = 6,
            xlEdgeBottom = 9,
            xlEdgeLeft = 7,
            xlEdgeRight = 10,
            xlEdgeTop = 8,
            xlInsideHorizontal = 12,
            xlInsideVertical = 11,
        }
        #endregion

        #region XlBorderWeight
        public enum XlBorderWeight
        {
            xlHairline = 1,
            xlMedium = -4138,
            xlThick = 4,
            xlThin = 2,
        }
        #endregion

        #region XlBuiltInDialog
        public enum XlBuiltInDialog
        {
            xlDialogActivate = 103,
            xlDialogActiveCellFont = 476,
            xlDialogAddChartAutoformat = 390,
            xlDialogAddinManager = 321,
            xlDialogAlignment = 43,
            xlDialogApplyNames = 133,
            xlDialogApplyStyle = 212,
            xlDialogAppMove = 170,
            xlDialogAppSize = 171,
            xlDialogArrangeAll = 12,
            xlDialogAssignToObject = 213,
            xlDialogAssignToTool = 293,
            xlDialogAttachText = 80,
            xlDialogAttachToolbars = 323,
            xlDialogAutoCorrect = 485,
            xlDialogAxes = 78,
            xlDialogBorder = 45,
            xlDialogCalculation = 32,
            xlDialogCellProtection = 46,
            xlDialogChangeLink = 166,
            xlDialogChartAddData = 392,
            xlDialogChartLocation = 527,
            xlDialogChartOptionsDataLabelMultiple = 724,
            xlDialogChartOptionsDataLabels = 505,
            xlDialogChartOptionsDataTable = 506,
            xlDialogChartSourceData = 540,
            xlDialogChartTrend = 350,
            xlDialogChartType = 526,
            xlDialogChartWizard = 288,
            xlDialogCheckboxProperties = 435,
            xlDialogClear = 52,
            xlDialogColorPalette = 161,
            xlDialogColumnWidth = 47,
            xlDialogCombination = 73,
            xlDialogConditionalFormatting = 583,
            xlDialogConsolidate = 191,
            xlDialogCopyChart = 147,
            xlDialogCopyPicture = 108,
            xlDialogCreateList = 796,
            xlDialogCreateNames = 62,
            xlDialogCreatePublisher = 217,
            xlDialogCreateRelationship = 1272,
            xlDialogCustomizeToolbar = 276,
            xlDialogCustomViews = 493,
            xlDialogDataDelete = 36,
            xlDialogDataLabel = 379,
            xlDialogDataLabelMultiple = 723,
            xlDialogDataSeries = 40,
            xlDialogDataValidation = 525,
            xlDialogDefineName = 61,
            xlDialogDefineStyle = 229,
            xlDialogDeleteFormat = 111,
            xlDialogDeleteName = 110,
            xlDialogDemote = 203,
            xlDialogDisplay = 27,
            xlDialogDocumentInspector = 862,
            xlDialogEditboxProperties = 438,
            xlDialogEditColor = 223,
            xlDialogEditDelete = 54,
            xlDialogEditionOptions = 251,
            xlDialogEditSeries = 228,
            xlDialogErrorbarX = 463,
            xlDialogErrorbarY = 464,
            xlDialogErrorChecking = 732,
            xlDialogEvaluateFormula = 709,
            xlDialogExternalDataProperties = 530,
            xlDialogExtract = 35,
            xlDialogFileDelete = 6,
            xlDialogFileSharing = 481,
            xlDialogFillGroup = 200,
            xlDialogFillWorkgroup = 301,
            xlDialogFilter = 447,
            xlDialogFilterAdvanced = 370,
            xlDialogFindFile = 475,
            xlDialogFont = 26,
            xlDialogFontProperties = 381,
            xlDialogFormatAuto = 269,
            xlDialogFormatChart = 465,
            xlDialogFormatCharttype = 423,
            xlDialogFormatFont = 150,
            xlDialogFormatLegend = 88,
            xlDialogFormatMain = 225,
            xlDialogFormatMove = 128,
            xlDialogFormatNumber = 42,
            xlDialogFormatOverlay = 226,
            xlDialogFormatSize = 129,
            xlDialogFormatText = 89,
            xlDialogFormulaFind = 64,
            xlDialogFormulaGoto = 63,
            xlDialogFormulaReplace = 130,
            xlDialogFunctionWizard = 450,
            xlDialogGallery3dArea = 193,
            xlDialogGallery3dBar = 272,
            xlDialogGallery3dColumn = 194,
            xlDialogGallery3dLine = 195,
            xlDialogGallery3dPie = 196,
            xlDialogGallery3dSurface = 273,
            xlDialogGalleryArea = 67,
            xlDialogGalleryBar = 68,
            xlDialogGalleryColumn = 69,
            xlDialogGalleryCustom = 388,
            xlDialogGalleryDoughnut = 344,
            xlDialogGalleryLine = 70,
            xlDialogGalleryPie = 71,
            xlDialogGalleryRadar = 249,
            xlDialogGalleryScatter = 72,
            xlDialogGoalSeek = 198,
            xlDialogGridlines = 76,
            xlDialogImportTextFile = 666,
            xlDialogInsert = 55,
            xlDialogInsertHyperlink = 596,
            xlDialogInsertObject = 259,
            xlDialogInsertPicture = 342,
            xlDialogInsertTitle = 380,
            xlDialogLabelProperties = 436,
            xlDialogListboxProperties = 437,
            xlDialogMacroOptions = 382,
            xlDialogMailEditMailer = 470,
            xlDialogMailLogon = 339,
            xlDialogMailNextLetter = 378,
            xlDialogMainChart = 85,
            xlDialogMainChartType = 185,
            xlDialogManageRelationships = 1271,
            xlDialogMenuEditor = 322,
            xlDialogMove = 262,
            xlDialogMyPermission = 834,
            xlDialogNameManager = 977,
            xlDialogNew = 119,
            xlDialogNewName = 978,
            xlDialogNewWebQuery = 667,
            xlDialogNote = 154,
            xlDialogObjectProperties = 207,
            xlDialogObjectProtection = 214,
            xlDialogOpen = 1,
            xlDialogOpenLinks = 2,
            xlDialogOpenMail = 188,
            xlDialogOpenText = 441,
            xlDialogOptionsCalculation = 318,
            xlDialogOptionsChart = 325,
            xlDialogOptionsEdit = 319,
            xlDialogOptionsGeneral = 356,
            xlDialogOptionsListsAdd = 458,
            xlDialogOptionsME = 647,
            xlDialogOptionsTransition = 355,
            xlDialogOptionsView = 320,
            xlDialogOutline = 142,
            xlDialogOverlay = 86,
            xlDialogOverlayChartType = 186,
            xlDialogPageSetup = 7,
            xlDialogParse = 91,
            xlDialogPasteNames = 58,
            xlDialogPasteSpecial = 53,
            xlDialogPatterns = 84,
            xlDialogPermission = 832,
            xlDialogPhonetic = 656,
            xlDialogPivotCalculatedField = 570,
            xlDialogPivotCalculatedItem = 572,
            xlDialogPivotClientServerSet = 689,
            xlDialogPivotFieldGroup = 433,
            xlDialogPivotFieldProperties = 313,
            xlDialogPivotFieldUngroup = 434,
            xlDialogPivotShowPages = 421,
            xlDialogPivotSolveOrder = 568,
            xlDialogPivotTableOptions = 567,
            xlDialogPivotTableSlicerConnections = 1183,
            xlDialogPivotTableWhatIfAnalysisSettings = 1153,
            xlDialogPivotTableWizard = 312,
            xlDialogPlacement = 300,
            xlDialogPrint = 8,
            xlDialogPrinterSetup = 9,
            xlDialogPrintPreview = 222,
            xlDialogPromote = 202,
            xlDialogProperties = 474,
            xlDialogPropertyFields = 754,
            xlDialogProtectDocument = 28,
            xlDialogProtectSharing = 620,
            xlDialogPublishAsWebPage = 653,
            xlDialogPushbuttonProperties = 445,
            xlDialogRecommendedPivotTables = 1258,
            xlDialogReplaceFont = 134,
            xlDialogRoutingSlip = 336,
            xlDialogRowHeight = 127,
            xlDialogRun = 17,
            xlDialogSaveAs = 5,
            xlDialogSaveCopyAs = 456,
            xlDialogSaveNewObject = 208,
            xlDialogSaveWorkbook = 145,
            xlDialogSaveWorkspace = 285,
            xlDialogScale = 87,
            xlDialogScenarioAdd = 307,
            xlDialogScenarioCells = 305,
            xlDialogScenarioEdit = 308,
            xlDialogScenarioMerge = 473,
            xlDialogScenarioSummary = 311,
            xlDialogScrollbarProperties = 420,
            xlDialogSearch = 731,
            xlDialogSelectSpecial = 132,
            xlDialogSendMail = 189,
            xlDialogSeriesAxes = 460,
            xlDialogSeriesOptions = 557,
            xlDialogSeriesOrder = 466,
            xlDialogSeriesShape = 504,
            xlDialogSeriesX = 461,
            xlDialogSeriesY = 462,
            xlDialogSetBackgroundPicture = 509,
            xlDialogSetManager = 1109,
            xlDialogSetMDXEditor = 1208,
            xlDialogSetPrintTitles = 23,
            xlDialogSetTupleEditorOnColumns = 1108,
            xlDialogSetTupleEditorOnRows = 1107,
            xlDialogSetUpdateStatus = 159,
            xlDialogShowDetail = 204,
            xlDialogShowToolbar = 220,
            xlDialogSize = 261,
            xlDialogSlicerCreation = 1182,
            xlDialogSlicerPivotTableConnections = 1184,
            xlDialogSlicerSettings = 1179,
            xlDialogSort = 39,
            xlDialogSortSpecial = 192,
            xlDialogSparklineInsertColumn = 1134,
            xlDialogSparklineInsertLine = 1133,
            xlDialogSparklineInsertWinLoss = 1135,
            xlDialogSplit = 137,
            xlDialogStandardFont = 190,
            xlDialogStandardWidth = 472,
            xlDialogStyle = 44,
            xlDialogSubscribeTo = 218,
            xlDialogSubtotalCreate = 398,
            xlDialogSummaryInfo = 474,
            xlDialogTable = 41,
            xlDialogTabOrder = 394,
            xlDialogTextToColumns = 422,
            xlDialogUnhide = 94,
            xlDialogUpdateLink = 201,
            xlDialogVbaInsertFile = 328,
            xlDialogVbaMakeAddin = 478,
            xlDialogVbaProcedureDefinition = 330,
            xlDialogView3d = 197,
            xlDialogWebOptionsBrowsers = 773,
            xlDialogWebOptionsEncoding = 686,
            xlDialogWebOptionsFiles = 684,
            xlDialogWebOptionsFonts = 687,
            xlDialogWebOptionsGeneral = 683,
            xlDialogWebOptionsPictures = 685,
            xlDialogWindowMove = 14,
            xlDialogWindowSize = 13,
            xlDialogWorkbookAdd = 281,
            xlDialogWorkbookCopy = 283,
            xlDialogWorkbookInsert = 354,
            xlDialogWorkbookMove = 282,
            xlDialogWorkbookName = 386,
            xlDialogWorkbookNew = 302,
            xlDialogWorkbookOptions = 284,
            xlDialogWorkbookProtect = 417,
            xlDialogWorkbookTabSplit = 415,
            xlDialogWorkbookUnhide = 384,
            xlDialogWorkgroup = 199,
            xlDialogWorkspace = 95,
            xlDialogZoom = 256,
        }
        #endregion

        #region XlCalcFor
        public enum XlCalcFor
        {
            xlAllValues = 0,
            xlColGroups = 2,
            xlRowGroups = 1,
        }
        #endregion

        #region XlCalcMemNumberFormatType
        public enum XlCalcMemNumberFormatType
        {
            xlNumberFormatTypeDefault = 0,
            xlNumberFormatTypeNumber = 1,
            xlNumberFormatTypePercent = 2,
        }
        #endregion

        #region XlCalculatedMemberType
        public enum XlCalculatedMemberType
        {
            xlCalculatedMeasure = 2,
            xlCalculatedMember = 0,
            xlCalculatedSet = 1,
        }
        #endregion

        #region XlCalculation
        public enum XlCalculation
        {
            xlCalculationAutomatic = -4105,
            xlCalculationManual = -4135,
            xlCalculationSemiautomatic = 2,
        }
        #endregion

        #region XlCalculationInterruptKey
        public enum XlCalculationInterruptKey
        {
            xlAnyKey = 2,
            xlEscKey = 1,
            xlNoKey = 0,
        }
        #endregion

        #region XlCalculationState
        public enum XlCalculationState
        {
            xlCalculating = 1,
            xlDone = 0,
            xlPending = 2,
        }
        #endregion

        #region XlCategoryLabelLevel
        public enum XlCategoryLabelLevel
        {
            xlCategoryLabelLevelAll = -1,
            xlCategoryLabelLevelCustom = -2,
            xlCategoryLabelLevelNone = -3,
        }
        #endregion

        #region XlCategoryType
        public enum XlCategoryType
        {
            xlAutomaticScale = -4105,
            xlCategoryScale = 2,
            xlTimeScale = 3,
        }
        #endregion

        #region XlCellChangedState
        public enum XlCellChangedState
        {
            xlCellChangeApplied = 3,
            xlCellChanged = 2,
            xlCellNotChanged = 1,
        }
        #endregion

        #region XlCellInsertionMode
        public enum XlCellInsertionMode
        {
            xlInsertDeleteCells = 1,
            xlInsertEntireRows = 2,
            xlOverwriteCells = 0,
        }
        #endregion

        #region XlCellType
        public enum XlCellType
        {
            xlCellTypeAllFormatConditions = -4172,
            xlCellTypeAllValidation = -4174,
            xlCellTypeBlanks = 4,
            xlCellTypeComments = -4144,
            xlCellTypeConstants = 2,
            xlCellTypeFormulas = -4123,
            xlCellTypeLastCell = 11,
            xlCellTypeSameFormatConditions = -4173,
            xlCellTypeSameValidation = -4175,
            xlCellTypeVisible = 12,
        }
        #endregion

        #region XlChartElement
        public enum XlChartElement
        {
            xlChartElementPositionAutomatic = -4105,
            xlChartElementPositionCustom = -4114,
        }
        #endregion

        #region XlChartGallery
        public enum XlChartGallery
        {
            xlAnyGallery = 23,
            xlBuiltIn = 21,
            xlUserDefined = 22,
        }
        #endregion

        #region XlChartItem
        public enum XlChartItem
        {
            xlAxis = 21,
            xlAxisTitle = 17,
            xlChartArea = 2,
            xlChartTitle = 4,
            xlCorners = 6,
            xlDataLabel = 0,
            xlDataTable = 7,
            xlDisplayUnitLabel = 30,
            xlDownBars = 20,
            xlDropLines = 26,
            xlErrorBars = 9,
            xlFloor = 23,
            xlHiLoLines = 25,
            xlLeaderLines = 29,
            xlLegend = 24,
            xlLegendEntry = 12,
            xlLegendKey = 13,
            xlMajorGridlines = 15,
            xlMinorGridlines = 16,
            xlNothing = 28,
            xlPivotChartDropZone = 32,
            xlPivotChartFieldButton = 31,
            xlPlotArea = 19,
            xlRadarAxisLabels = 27,
            xlSeries = 3,
            xlSeriesLines = 22,
            xlShape = 14,
            xlTrendline = 8,
            xlUpBars = 18,
            xlWalls = 5,
            xlXErrorBars = 10,
            xlYErrorBars = 11,
        }
        #endregion

        #region XlChartLocation
        public enum XlChartLocation
        {
            xlLocationAsNewSheet = 1,
            xlLocationAsObject = 2,
            xlLocationAutomatic = 3,
        }
        #endregion

        #region XlChartPicturePlacement
        public enum XlChartPicturePlacement
        {
            xlAllFaces = 7,
            xlEnd = 2,
            xlEndSides = 3,
            xlFront = 4,
            xlFrontEnd = 6,
            xlFrontSides = 5,
            xlSides = 1,
        }
        #endregion

        #region XlChartPictureType
        public enum XlChartPictureType
        {
            xlStack = 2,
            xlStackScale = 3,
            xlStretch = 1,
        }
        #endregion

        #region XlChartSplitType
        public enum XlChartSplitType
        {
            xlSplitByCustomSplit = 4,
            xlSplitByPercentValue = 3,
            xlSplitByPosition = 1,
            xlSplitByValue = 2,
        }
        #endregion

        #region XlChartType
        public enum XlChartType
        {
            xl3DArea = -4098,
            xl3DAreaStacked = 78,
            xl3DAreaStacked100 = 79,
            xl3DBarClustered = 60,
            xl3DBarStacked = 61,
            xl3DBarStacked100 = 62,
            xl3DColumn = -4100,
            xl3DColumnClustered = 54,
            xl3DColumnStacked = 55,
            xl3DColumnStacked100 = 56,
            xl3DLine = -4101,
            xl3DPie = -4102,
            xl3DPieExploded = 70,
            xlArea = 1,
            xlAreaStacked = 76,
            xlAreaStacked100 = 77,
            xlBarClustered = 57,
            xlBarOfPie = 71,
            xlBarStacked = 58,
            xlBarStacked100 = 59,
            xlBubble = 15,
            xlBubble3DEffect = 87,
            xlColumnClustered = 51,
            xlColumnStacked = 52,
            xlColumnStacked100 = 53,
            xlConeBarClustered = 102,
            xlConeBarStacked = 103,
            xlConeBarStacked100 = 104,
            xlConeCol = 105,
            xlConeColClustered = 99,
            xlConeColStacked = 100,
            xlConeColStacked100 = 101,
            xlCylinderBarClustered = 95,
            xlCylinderBarStacked = 96,
            xlCylinderBarStacked100 = 97,
            xlCylinderCol = 98,
            xlCylinderColClustered = 92,
            xlCylinderColStacked = 93,
            xlCylinderColStacked100 = 94,
            xlDoughnut = -4120,
            xlDoughnutExploded = 80,
            xlLine = 4,
            xlLineMarkers = 65,
            xlLineMarkersStacked = 66,
            xlLineMarkersStacked100 = 67,
            xlLineStacked = 63,
            xlLineStacked100 = 64,
            xlPie = 5,
            xlPieExploded = 69,
            xlPieOfPie = 68,
            xlPyramidBarClustered = 109,
            xlPyramidBarStacked = 110,
            xlPyramidBarStacked100 = 111,
            xlPyramidCol = 112,
            xlPyramidColClustered = 106,
            xlPyramidColStacked = 107,
            xlPyramidColStacked100 = 108,
            xlRadar = -4151,
            xlRadarFilled = 82,
            xlRadarMarkers = 81,
            xlStockHLC = 88,
            xlStockOHLC = 89,
            xlStockVHLC = 90,
            xlStockVOHLC = 91,
            xlSurface = 83,
            xlSurfaceTopView = 85,
            xlSurfaceTopViewWireframe = 86,
            xlSurfaceWireframe = 84,
            xlXYScatter = -4169,
            xlXYScatterLines = 74,
            xlXYScatterLinesNoMarkers = 75,
            xlXYScatterSmooth = 72,
            xlXYScatterSmoothNoMarkers = 73,
        }
        #endregion

        #region XlCheckInVersionType
        public enum XlCheckInVersionType
        {
            xlCheckInMajorVersion = 1,
            xlCheckInMinorVersion = 0,
            xlCheckInOverwriteVersion = 2,
        }
        #endregion

        #region XlClipboardFormat
        public enum XlClipboardFormat
        {
            xlClipboardFormatBIFF = 8,
            xlClipboardFormatBIFF2 = 18,
            xlClipboardFormatBIFF3 = 20,
            xlClipboardFormatBIFF4 = 30,
            xlClipboardFormatBinary = 15,
            xlClipboardFormatBitmap = 9,
            xlClipboardFormatCGM = 13,
            xlClipboardFormatCSV = 5,
            xlClipboardFormatDIF = 4,
            xlClipboardFormatDspText = 12,
            xlClipboardFormatEmbeddedObject = 21,
            xlClipboardFormatEmbedSource = 22,
            xlClipboardFormatLink = 11,
            xlClipboardFormatLinkSource = 23,
            xlClipboardFormatLinkSourceDesc = 32,
            xlClipboardFormatMovie = 24,
            xlClipboardFormatNative = 14,
            xlClipboardFormatObjectDesc = 31,
            xlClipboardFormatObjectLink = 19,
            xlClipboardFormatOwnerLink = 17,
            xlClipboardFormatPICT = 2,
            xlClipboardFormatPrintPICT = 3,
            xlClipboardFormatRTF = 7,
            xlClipboardFormatScreenPICT = 29,
            xlClipboardFormatStandardFont = 28,
            xlClipboardFormatStandardScale = 27,
            xlClipboardFormatSYLK = 6,
            xlClipboardFormatTable = 16,
            xlClipboardFormatText = 0,
            xlClipboardFormatToolFace = 25,
            xlClipboardFormatToolFacePICT = 26,
            xlClipboardFormatVALU = 1,
            xlClipboardFormatWK1 = 10,
        }
        #endregion

        #region XlCmdType
        public enum XlCmdType
        {
            xlCmdCube = 1,
            xlCmdDAX = 8,
            xlCmdDefault = 4,
            xlCmdExcel = 7,
            xlCmdList = 5,
            xlCmdSql = 2,
            xlCmdTable = 3,
            xlCmdTableCollection = 6,
        }
        #endregion

        #region XlColorIndex
        public enum XlColorIndex
        {
            xlColorIndexAutomatic = -4105,
            xlColorIndexNone = -4142,
        }
        #endregion

        #region XlColumnDataType
        public enum XlColumnDataType
        {
            xlDMYFormat = 4,
            xlDYMFormat = 7,
            xlEMDFormat = 10,
            xlGeneralFormat = 1,
            xlMDYFormat = 3,
            xlMYDFormat = 6,
            xlSkipColumn = 9,
            xlTextFormat = 2,
            xlYDMFormat = 8,
            xlYMDFormat = 5,
        }
        #endregion

        #region XlCommandUnderlines
        public enum XlCommandUnderlines
        {
            xlCommandUnderlinesAutomatic = -4105,
            xlCommandUnderlinesOff = -4146,
            xlCommandUnderlinesOn = 1,
        }
        #endregion

        #region XlCommentDisplayMode
        public enum XlCommentDisplayMode
        {
            xlCommentAndIndicator = 1,
            xlCommentIndicatorOnly = -1,
            xlNoIndicator = 0,
        }
        #endregion

        #region XlConditionValueTypes
        public enum XlConditionValueTypes
        {
            xlConditionValueAutomaticMax = 7,
            xlConditionValueAutomaticMin = 6,
            xlConditionValueFormula = 4,
            xlConditionValueHighestValue = 2,
            xlConditionValueLowestValue = 1,
            xlConditionValueNone = -1,
            xlConditionValueNumber = 0,
            xlConditionValuePercent = 3,
            xlConditionValuePercentile = 5,
        }
        #endregion

        #region XlConnectionType
        public enum XlConnectionType
        {
            xlConnectionTypeDATAFEED = 6,
            xlConnectionTypeMODEL = 7,
            xlConnectionTypeODBC = 2,
            xlConnectionTypeOLEDB = 1,
            xlConnectionTypeTEXT = 4,
            xlConnectionTypeWEB = 5,
            xlConnectionTypeWORKSHEET = 8,
            xlConnectionTypeXMLMAP = 3,
        }
        #endregion

        #region XlConsolidationFunction
        public enum XlConsolidationFunction
        {
            xlAverage = -4106,
            xlCount = -4112,
            xlCountNums = -4113,
            xlmax = -4136,
            xlMin = -4139,
            xlProduct = -4149,
            xlStDev = -4155,
            xlStDevP = -4156,
            xlSum = -4157,
            xlUnknown = 1000,
            xlVar = -4164,
            xlVarP = -4165,
        }
        #endregion

        #region XlContainsOperator
        public enum XlContainsOperator
        {
            xlBeginsWith = 2,
            xlContains = 0,
            xlDoesNotContain = 1,
            xlEndsWith = 3,
        }
        #endregion

        #region XlCopyPictureFormat
        public enum XlCopyPictureFormat
        {
            xlBitmap = 2,
            xlPicture = -4147,
        }
        #endregion

        #region XlCorruptLoad
        public enum XlCorruptLoad
        {
            xlExtractData = 2,
            xlNormalLoad = 0,
            xlRepairFile = 1,
        }
        #endregion

        #region XlCreator
        public enum XlCreator
        {
            xlCreatorCode = 1480803660,
        }
        #endregion

        #region XlCredentialsMethod
        public enum XlCredentialsMethod
        {
            CredentialsMethodIntegrated = 0,
            CredentialsMethodNone = 1,
            CredentialsMethodStored = 2,
        }
        #endregion

        #region XlCubeFieldSubType
        public enum XlCubeFieldSubType
        {
            xlCubeAttribute = 4,
            xlCubeCalculatedMeasure = 5,
            xlCubeHierarchy = 1,
            xlCubeImplicitMeasure = 11,
            xlCubeKPIGoal = 7,
            xlCubeKPIStatus = 8,
            xlCubeKPITrend = 9,
            xlCubeKPIValue = 6,
            xlCubeKPIWeight = 10,
            xlCubeMeasure = 2,
            xlCubeSet = 3,
        }
        #endregion

        #region XlCubeFieldType
        public enum XlCubeFieldType
        {
            xlHierarchy = 1,
            xlMeasure = 2,
            xlSet = 3,
        }
        #endregion

        #region XlCutCopyMode
        public enum XlCutCopyMode
        {
            xlCopy = 1,
            xlCut = 2,
        }
        #endregion

        #region XlCVError
        public enum XlCVError
        {
            xlErrDiv0 = 2007,
            xlErrNA = 2042,
            xlErrName = 2029,
            xlErrNull = 2000,
            xlErrNum = 2036,
            xlErrRef = 2023,
            xlErrValue = 2015,
        }
        #endregion

        #region XlDataBarAxisPosition
        public enum XlDataBarAxisPosition
        {
            xlDataBarAxisAutomatic = 0,
            xlDataBarAxisMidpoint = 1,
            xlDataBarAxisNone = 2,
        }
        #endregion

        #region XlDataBarBorderType
        public enum XlDataBarBorderType
        {
            xlDataBarBorderNone = 0,
            xlDataBarBorderSolid = 1,
        }
        #endregion

        #region XlDataBarFillType
        public enum XlDataBarFillType
        {
            xlDataBarFillGradient = 1,
            xlDataBarFillSolid = 0,
        }
        #endregion

        #region XlDataBarNegativeColorType
        public enum XlDataBarNegativeColorType
        {
            xlDataBarColor = 0,
            xlDataBarSameAsPositive = 1,
        }
        #endregion

        #region XlDataLabelPosition
        public enum XlDataLabelPosition
        {
            xlLabelPositionAbove = 0,
            xlLabelPositionBelow = 1,
            xlLabelPositionBestFit = 5,
            xlLabelPositionCenter = -4108,
            xlLabelPositionCustom = 7,
            xlLabelPositionInsideBase = 4,
            xlLabelPositionInsideEnd = 3,
            xlLabelPositionLeft = -4131,
            xlLabelPositionMixed = 6,
            xlLabelPositionOutsideEnd = 2,
            xlLabelPositionRight = -4152,
        }
        #endregion

        #region XlDataLabelSeparator
        public enum XlDataLabelSeparator
        {
            xlDataLabelSeparatorDefault = 1,
        }
        #endregion

        #region XlDataLabelsType
        public enum XlDataLabelsType
        {
            xlDataLabelsShowBubbleSizes = 6,
            xlDataLabelsShowLabel = 4,
            xlDataLabelsShowLabelAndPercent = 5,
            xlDataLabelsShowNone = -4142,
            xlDataLabelsShowPercent = 3,
            xlDataLabelsShowValue = 2,
        }
        #endregion

        #region XlDataSeriesDate
        public enum XlDataSeriesDate
        {
            xlDay = 1,
            xlMonth = 3,
            xlWeekday = 2,
            xlYear = 4,
        }
        #endregion

        #region XlDataSeriesType
        public enum XlDataSeriesType
        {
            xlAutoFill = 4,
            xlChronological = 3,
            xlDataSeriesLinear = -4132,
            xlGrowth = 2,
        }
        #endregion

        #region XlDeleteShiftDirection
        public enum XlDeleteShiftDirection
        {
            xlShiftToLeft = -4159,
            xlShiftUp = -4162,
        }
        #endregion

        #region XlDirection
        public enum XlDirection
        {
            xlDown = -4121,
            xlToLeft = -4159,
            xlToRight = -4161,
            xlUp = -4162,
        }
        #endregion

        #region XlDisplayBlanksAs
        public enum XlDisplayBlanksAs
        {
            xlInterpolated = 3,
            xlNotPlotted = 1,
            xlZero = 2,
        }
        #endregion

        #region XlDisplayDrawingObjects
        public enum XlDisplayDrawingObjects
        {
            xlDisplayShapes = -4104,
            xlHide = 3,
            xlPlaceholders = 2,
        }
        #endregion

        #region XlDisplayUnit
        public enum XlDisplayUnit
        {
            xlHundredMillions = -8,
            xlHundreds = -2,
            xlHundredThousands = -5,
            xlMillionMillions = -10,
            xlMillions = -6,
            xlTenMillions = -7,
            xlTenThousands = -4,
            xlThousandMillions = -9,
            xlThousands = -3,
        }
        #endregion

        #region XlDupeUnique
        public enum XlDupeUnique
        {
            xlDuplicate = 1,
            xlUnique = 0,
        }
        #endregion

        #region XlDVAlertStyle
        public enum XlDVAlertStyle
        {
            xlValidAlertInformation = 3,
            xlValidAlertStop = 1,
            xlValidAlertWarning = 2,
        }
        #endregion

        #region XlDVType
        public enum XlDVType
        {
            xlValidateCustom = 7,
            xlValidateDate = 4,
            xlValidateDecimal = 2,
            xlValidateInputOnly = 0,
            xlValidateList = 3,
            xlValidateTextLength = 6,
            xlValidateTime = 5,
            xlValidateWholeNumber = 1,
        }
        #endregion

        #region XlDynamicFilterCriteria
        public enum XlDynamicFilterCriteria
        {
            xlFilterAboveAverage = 33,
            xlFilterAllDatesInPeriodApril = 24,
            xlFilterAllDatesInPeriodAugust = 28,
            xlFilterAllDatesInPeriodDecember = 32,
            xlFilterAllDatesInPeriodFebruray = 22,
            xlFilterAllDatesInPeriodJanuary = 21,
            xlFilterAllDatesInPeriodJuly = 27,
            xlFilterAllDatesInPeriodJune = 26,
            xlFilterAllDatesInPeriodMarch = 23,
            xlFilterAllDatesInPeriodMay = 25,
            xlFilterAllDatesInPeriodNovember = 31,
            xlFilterAllDatesInPeriodOctober = 30,
            xlFilterAllDatesInPeriodQuarter1 = 17,
            xlFilterAllDatesInPeriodQuarter2 = 18,
            xlFilterAllDatesInPeriodQuarter3 = 19,
            xlFilterAllDatesInPeriodQuarter4 = 20,
            xlFilterAllDatesInPeriodSeptember = 29,
            xlFilterBelowAverage = 34,
            xlFilterLastMonth = 8,
            xlFilterLastQuarter = 11,
            xlFilterLastWeek = 5,
            xlFilterLastYear = 14,
            xlFilterNextMonth = 9,
            xlFilterNextQuarter = 12,
            xlFilterNextWeek = 6,
            xlFilterNextYear = 15,
            xlFilterThisMonth = 7,
            xlFilterThisQuarter = 10,
            xlFilterThisWeek = 4,
            xlFilterThisYear = 13,
            xlFilterToday = 1,
            xlFilterTomorrow = 3,
            xlFilterYearToDate = 16,
            xlFilterYesterday = 2,
        }
        #endregion

        #region XlEditionFormat
        public enum XlEditionFormat
        {
            xlBIFF = 2,
            xlPICT = 1,
            xlRTF = 4,
            xlVALU = 8,
        }
        #endregion

        #region XlEditionOptionsOption
        public enum XlEditionOptionsOption
        {
            xlAutomaticUpdate = 4,
            xlCancel = 1,
            xlChangeAttributes = 6,
            xlManualUpdate = 5,
            xlOpenSource = 3,
            xlSelect = 3,
            xlSendPublisher = 2,
            xlUpdateSubscriber = 2,
        }
        #endregion

        #region XlEditionType
        public enum XlEditionType
        {
            xlPublisher = 1,
            xlSubscriber = 2,
        }
        #endregion

        #region XlEnableCancelKey
        public enum XlEnableCancelKey
        {
            xlDisabled = 0,
            xlErrorHandler = 2,
            xlInterrupt = 1,
        }
        #endregion

        #region XlEnableSelection
        public enum XlEnableSelection
        {
            xlNoRestrictions = 0,
            xlNoSelection = -4142,
            xlUnlockedCells = 1,
        }
        #endregion

        #region XlEndStyleCap
        public enum XlEndStyleCap
        {
            xlCap = 1,
            xlNoCap = 2,
        }
        #endregion

        #region XlErrorBarDirection
        public enum XlErrorBarDirection
        {
            xlX = -4168,
            xlY = 1,
        }
        #endregion

        #region XlErrorBarInclude
        public enum XlErrorBarInclude
        {
            xlErrorBarIncludeBoth = 1,
            xlErrorBarIncludeMinusValues = 3,
            xlErrorBarIncludeNone = -4142,
            xlErrorBarIncludePlusValues = 2,
        }
        #endregion

        #region XlErrorBarType
        public enum XlErrorBarType
        {
            xlErrorBarTypeCustom = -4114,
            xlErrorBarTypeFixedValue = 1,
            xlErrorBarTypePercent = 2,
            xlErrorBarTypeStDev = -4155,
            xlErrorBarTypeStError = 4,
        }
        #endregion

        #region XlErrorChecks
        public enum XlErrorChecks
        {
            xlEmptyCellReferences = 7,
            xlEvaluateToError = 1,
            xlInconsistentFormula = 4,
            xlListDataValidation = 8,
            xlNumberAsText = 3,
            xlOmittedCells = 5,
            xlTextDate = 2,
            xlUnlockedFormulaCells = 6,
        }
        #endregion

        #region XlFileAccess
        public enum XlFileAccess
        {
            xlReadOnly = 3,
            xlReadWrite = 2,
        }
        #endregion

        #region XlFileFormat
        public enum XlFileFormat
        {
            xlAddIn = 18,
            xlAddIn8 = 18,
            xlCSV = 6,
            xlCSVMac = 22,
            xlCSVMSDOS = 24,
            xlCSVWindows = 23,
            xlCurrentPlatformText = -4158,
            xlDBF2 = 7,
            xlDBF3 = 8,
            xlDBF4 = 11,
            xlDIF = 9,
            xlExcel12 = 50,
            xlExcel2 = 16,
            xlExcel2FarEast = 27,
            xlExcel3 = 29,
            xlExcel4 = 33,
            xlExcel4Workbook = 35,
            xlExcel5 = 39,
            xlExcel7 = 39,
            xlExcel8 = 56,
            xlExcel9795 = 43,
            xlHtml = 44,
            xlIntlAddIn = 26,
            xlIntlMacro = 25,
            xlOpenDocumentSpreadsheet = 60,
            xlOpenXMLAddIn = 55,
            xlOpenXMLStrictWorkbook = 61,
            xlOpenXMLTemplate = 54,
            xlOpenXMLTemplateMacroEnabled = 53,
            xlOpenXMLWorkbook = 51,
            xlOpenXMLWorkbookMacroEnabled = 52,
            xlSYLK = 2,
            xlTemplate = 17,
            xlTemplate8 = 17,
            xlTextMac = 19,
            xlTextMSDOS = 21,
            xlTextPrinter = 36,
            xlTextWindows = 20,
            xlUnicodeText = 42,
            xlWebArchive = 45,
            xlWJ2WD1 = 14,
            xlWJ3 = 40,
            xlWJ3FJ3 = 41,
            xlWK1 = 5,
            xlWK1ALL = 31,
            xlWK1FMT = 30,
            xlWK3 = 15,
            xlWK3FM3 = 32,
            xlWK4 = 38,
            xlWKS = 4,
            xlWorkbookDefault = 51,
            xlWorkbookNormal = -4143,
            xlWorks2FarEast = 28,
            xlWQ1 = 34,
            xlXMLSpreadsheet = 46,
        }
        #endregion

        #region XlFileValidationPivotMode
        public enum XlFileValidationPivotMode
        {
            xlFileValidationPivotDefault = 0,
            xlFileValidationPivotRun = 1,
            xlFileValidationPivotSkip = 2,
        }
        #endregion

        #region XlFillWith
        public enum XlFillWith
        {
            xlFillWithAll = -4104,
            xlFillWithContents = 2,
            xlFillWithFormats = -4122,
        }
        #endregion

        #region XlFilterAction
        public enum XlFilterAction
        {
            xlFilterCopy = 2,
            xlFilterInPlace = 1,
        }
        #endregion

        #region XlFilterAllDatesInPeriod
        public enum XlFilterAllDatesInPeriod
        {
            xlFilterAllDatesInPeriodDay = 2,
            xlFilterAllDatesInPeriodHour = 3,
            xlFilterAllDatesInPeriodMinute = 4,
            xlFilterAllDatesInPeriodMonth = 1,
            xlFilterAllDatesInPeriodSecond = 5,
            xlFilterAllDatesInPeriodYear = 0,
        }
        #endregion

        #region XlFilterStatus
        public enum XlFilterStatus
        {
            xlFilterStatusOK = 0,
            xlFilterStatusDateWrongOrder = 1,
            xlFilterStatusDateHasTime = 2,
            xlFilterStatusInvalidDate = 3,
        }
        #endregion

        #region XlFindLookIn
        public enum XlFindLookIn
        {
            xlComments = -4144,
            xlFormulas = -4123,
            xlValues = -4163,
        }
        #endregion

        #region XlFixedFormatQuality
        public enum XlFixedFormatQuality
        {
            xlQualityMinimum = 1,
            xlQualityStandard = 0,
        }
        #endregion

        #region XlFixedFormatType
        public enum XlFixedFormatType
        {
            xlTypePDF = 0,
            xlTypeXPS = 0,
        }
        #endregion

        #region XlFormatConditionOperator
        public enum XlFormatConditionOperator
        {
            xlBetween = 1,
            xlEqual = 3,
            xlGreater = 5,
            xlGreaterEqual = 7,
            xlLess = 6,
            xlLessEqual = 8,
            xlNotBetween = 2,
            xlNotEqual = 4,
        }
        #endregion

        #region XlFormatConditionType
        public enum XlFormatConditionType
        {
            xlAboveAverageCondition = 12,
            xlBlanksCondition = 10,
            xlCellValue = 1,
            xlColorScale = 3,
            xlDatabar = 4,
            xlErrorsCondition = 16,
            xlExpression = 2,
            XlIconSet = 6,
            xlNoBlanksCondition = 13,
            xlNoErrorsCondition = 17,
            xlTextString = 9,
            xlTimePeriod = 11,
            xlTop10 = 5,
            xlUniqueValues = 8,
        }
        #endregion

        #region XlFormatFilterTypes
        public enum XlFormatFilterTypes
        {
            FilterBottom = 0,
            FilterBottomPercent = 2,
            FilterTop = 1,
            FilterTopPercent = 3,
        }
        #endregion

        #region XlFormControl
        public enum XlFormControl
        {
            xlButtonControl = 0,
            xlCheckBox = 1,
            xlDropDown = 2,
            xlEditBox = 3,
            xlGroupBox = 4,
            xlLabel = 5,
            xlListBox = 6,
            xlOptionButton = 7,
            xlScrollBar = 8,
            xlSpinner = 9,
        }
        #endregion

        #region XlFormulaLabel
        public enum XlFormulaLabel
        {
            xlColumnLabels = 2,
            xlMixedLabels = 3,
            xlNoLabels = -4142,
            xlRowLabels = 1,
        }
        #endregion

        #region XlGenerateTableRefs
        public enum XlGenerateTableRefs
        {
            xlA1TableRefs = 0,
            xlTableNames = 1,
        }
        #endregion

        #region XlGradientFillType
        public enum XlGradientFillType
        {
            GradientFillLinear = 0,
            GradientFillPath = 1,
        }
        #endregion

        #region XlHAlign
        public enum XlHAlign
        {
            xlHAlignCenter = -4108,
            xlHAlignCenterAcrossSelection = 7,
            xlHAlignDistributed = -4117,
            xlHAlignFill = 5,
            xlHAlignGeneral = 1,
            xlHAlignJustify = -4130,
            xlHAlignLeft = -4131,
            xlHAlignRight = -4152,
        }
        #endregion

        #region XlHebrewModes
        public enum XlHebrewModes
        {
            xlHebrewFullScript = 0,
            xlHebrewMixedAuthorizedScript = 3,
            xlHebrewMixedScript = 2,
            xlHebrewPartialScript = 1,
        }
        #endregion

        #region XlHighlightChangesTime
        public enum XlHighlightChangesTime
        {
            xlAllChanges = 2,
            xlNotYetReviewed = 3,
            xlSinceMyLastSave = 1,
        }
        #endregion

        #region XlHtmlType
        public enum XlHtmlType
        {
            xlHtmlCalc = 1,
            xlHtmlChart = 3,
            xlHtmlList = 2,
            xlHtmlStatic = 0,
        }
        #endregion

        #region XlIcon
        public enum XlIcon
        {
            xlIcon0Bars = 37,
            xlIcon0FilledBoxes = 52,
            xlIcon1Bar = 38,
            xlIcon1FilledBox = 51,
            xlIcon2Bars = 39,
            xlIcon2FilledBoxes = 50,
            xlIcon3Bars = 40,
            xlIcon3FilledBoxes = 49,
            xlIcon4Bars = 41,
            xlIcon4FilledBoxes = 48,
            xlIconBlackCircle = 32,
            xlIconBlackCircleWithBorder = 13,
            xlIconCircleWithOneWhiteQuarter = 33,
            xlIconCircleWithThreeWhiteQuarters = 35,
            xlIconCircleWithTwoWhiteQuarters = 34,
            xlIconGoldStar = 42,
            xlIconGrayCircle = 31,
            xlIconGrayDownArrow = 6,
            xlIconGrayDownInclineArrow = 28,
            xlIconGraySideArrow = 5,
            xlIconGrayUpArrow = 4,
            xlIconGrayUpInclineArrow = 27,
            xlIconGreenCheck = 22,
            xlIconGreenCheckSymbol = 19,
            xlIconGreenCircle = 10,
            xlIconGreenFlag = 7,
            xlIconGreenTrafficLight = 14,
            xlIconGreenUpArrow = 1,
            xlIconGreenUpTriangle = 45,
            xlIconHalfGoldStar = 43,
            xlIconNoCellIcon = -1,
            xlIconPinkCircle = 30,
            xlIconRedCircle = 29,
            xlIconRedCircleWithBorder = 12,
            xlIconRedCross = 24,
            xlIconRedCrossSymbol = 21,
            xlIconRedDiamond = 18,
            xlIconRedDownArrow = 3,
            xlIconRedDownTriangle = 47,
            xlIconRedFlag = 9,
            xlIconRedTrafficLight = 16,
            xlIconSilverStar = 44,
            xlIconWhiteCircleAllWhiteQuarters = 36,
            xlIconYellowCircle = 11,
            xlIconYellowDash = 46,
            xlIconYellowDownInclineArrow = 26,
            xlIconYellowExclamation = 23,
            xlIconYellowExclamationSymbol = 20,
            xlIconYellowFlag = 8,
            xlIconYellowSideArrow = 2,
            xlIconYellowTrafficLight = 15,
            xlIconYellowTriangle = 17,
            xlIconYellowUpInclineArrow = 25,
        }
        #endregion

        #region XlIconSet
        public enum XlIconSet
        {
            xl3Arrows = 1,
            xl3ArrowsGray = 2,
            xl3Flags = 3,
            xl3Signs = 6,
            xl3Symbols = 7,
            xl3TrafficLights1 = 4,
            xl3TrafficLights2 = 5,
            xl4Arrows = 8,
            xl4ArrowsGray = 9,
            xl4CRV = 11,
            xl4RedToBlack = 10,
            xl4TrafficLights = 12,
            xl5Arrows = 13,
            xl5ArrowsGray = 14,
            xl5CRV = 15,
            xl5Quarters = 16,
        }
        #endregion

        #region XlIMEMode
        public enum XlIMEMode
        {
            xlIMEModeAlpha = 8,
            xlIMEModeAlphaFull = 7,
            xlIMEModeDisable = 3,
            xlIMEModeHangul = 10,
            xlIMEModeHangulFull = 9,
            xlIMEModeHiragana = 4,
            xlIMEModeKatakana = 5,
            xlIMEModeKatakanaHalf = 6,
            xlIMEModeNoControl = 0,
            xlIMEModeOff = 2,
            xlIMEModeOn = 1,
        }
        #endregion

        #region XlImportDataAs
        public enum XlImportDataAs
        {
            xlPivotTableReport = 1,
            xlQueryTable = 0,
        }
        #endregion

        #region XlInsertFormatOrigin
        public enum XlInsertFormatOrigin
        {
            xlFormatFromLeftOrAbove = 0,
            xlFormatFromRightOrBelow = 1,
        }
        #endregion

        #region XlInsertShiftDirection
        public enum XlInsertShiftDirection
        {
            xlShiftDown = -4121,
            xlShiftToRight = -4161,
        }
        #endregion

        #region XlLayoutFormType
        public enum XlLayoutFormType
        {
            xlOutline = 1,
            xlTabular = 0,
        }
        #endregion

        #region XlLayoutRowType
        public enum XlLayoutRowType
        {
            xlCompactRow = 0,
            xlOutlineRow = 2,
            xlTabularRow = 1,
        }
        #endregion

        #region XlLegendPosition
        public enum XlLegendPosition
        {
            xlLegendPositionBottom = -4107,
            xlLegendPositionCorner = 2,
            xlLegendPositionLeft = -4131,
            xlLegendPositionRight = -4152,
            xlLegendPositionTop = -4160,
        }
        #endregion

        #region XlLineStyle
        public enum XlLineStyle
        {
            xlContinuous = 1,
            xlDash = -4115,
            xlDashDot = 4,
            xlDashDotDot = 5,
            xlDot = -4118,
            xlDouble = -4119,
            xlLineStyleNone = -4142,
            xlSlantDashDot = 13,
        }
        #endregion

        #region XlLink
        public enum XlLink
        {
            xlExcelLinks = 1,
            xlOLELinks = 2,
            xlPublishers = 5,
            xlSubscribers = 6,
        }
        #endregion

        #region XlLinkInfo
        public enum XlLinkInfo
        {
            xlEditionDate = 2,
            xlLinkInfoStatus = 3,
            xlUpdateState = 1,
        }
        #endregion

        #region XlLinkInfoType
        public enum XlLinkInfoType
        {
            xlLinkInfoOLELinks = 2,
            xlLinkInfoPublishers = 5,
            xlLinkInfoSubscribers = 6,
        }
        #endregion

        #region XlLinkStatus
        public enum XlLinkStatus
        {
            xlLinkStatusCopiedValues = 10,
            xlLinkStatusIndeterminate = 5,
            xlLinkStatusInvalidName = 7,
            xlLinkStatusMissingFile = 1,
            xlLinkStatusMissingSheet = 2,
            xlLinkStatusNotStarted = 6,
            xlLinkStatusOK = 0,
            xlLinkStatusOld = 3,
            xlLinkStatusSourceNotCalculated = 4,
            xlLinkStatusSourceNotOpen = 8,
            xlLinkStatusSourceOpen = 9,
        }
        #endregion

        #region XlLinkType
        public enum XlLinkType
        {
            xlLinkTypeExcelLinks = 1,
            xlLinkTypeOLELinks = 2,
        }
        #endregion

        #region XlListConflict
        public enum XlListConflict
        {
            xlListConflictDialog = 0,
            xlListConflictDiscardAllConflicts = 2,
            xlListConflictError = 3,
            xlListConflictRetryAllConflicts = 1,
        }
        #endregion

        #region XlListDataType
        public enum XlListDataType
        {
            xlListDataTypeCheckbox = 9,
            xlListDataTypeChoice = 6,
            xlListDataTypeChoiceMulti = 7,
            xlListDataTypeCounter = 11,
            xlListDataTypeCurrency = 4,
            xlListDataTypeDateTime = 5,
            xlListDataTypeHyperLink = 10,
            xlListDataTypeListLookup = 8,
            xlListDataTypeMultiLineRichText = 12,
            xlListDataTypeMultiLineText = 2,
            xlListDataTypeNone = 0,
            xlListDataTypeNumber = 3,
            xlListDataTypeText = 1,
        }
        #endregion

        #region XlListObjectSourceType
        public enum XlListObjectSourceType
        {
            xlSrcExternal = 0,
            xlSrcModel = 4,
            xlSrcQuery = 3,
            xlSrcRange = 1,
            xlSrcXml = 2,
        }
        #endregion

        #region XlLocationInTable
        public enum XlLocationInTable
        {
            xlColumnHeader = -4110,
            xlColumnItem = 5,
            xlDataHeader = 3,
            xlDataItem = 7,
            xlPageHeader = 2,
            xlPageItem = 6,
            xlRowHeader = -4153,
            xlRowItem = 4,
            xlTableBody = 8,
        }
        #endregion

        #region XlLookAt
        public enum XlLookAt
        {
            xlPart = 2,
            xlWhole = 1,
        }
        #endregion

        #region XlLookFor
        public enum XlLookFor
        {
            LookForBlanks = 0,
            LookForErrors = 1,
            LookForFormulas = 2,
        }
        #endregion

        #region XlMailSystem
        public enum XlMailSystem
        {
            xlMAPI = 1,
            xlNoMailSystem = 0,
            xlPowerTalk = 2,
        }
        #endregion

        #region XlMarkerStyle
        public enum XlMarkerStyle
        {
            xlMarkerStyleAutomatic = -4105,
            xlMarkerStyleCircle = 8,
            xlMarkerStyleDash = -4115,
            xlMarkerStyleDiamond = 2,
            xlMarkerStyleDot = -4118,
            xlMarkerStyleNone = -4142,
            xlMarkerStylePicture = -4147,
            xlMarkerStylePlus = 9,
            xlMarkerStyleSquare = 1,
            xlMarkerStyleStar = 5,
            xlMarkerStyleTriangle = 3,
            xlMarkerStyleX = -4168,
        }
        #endregion

        #region XlMeasurementUnits
        public enum XlMeasurementUnits
        {
            xlCentimeters = 1,
            xlInches = 0,
            xlMillimeters = 2,
        }
        #endregion

        #region XlMouseButton
        public enum XlMouseButton
        {
            xlNoButton = 0,
            xlPrimaryButton = 1,
            xlSecondaryButton = 2,
        }
        #endregion

        #region XlMousePointer
        public enum XlMousePointer
        {
            xlDefault = -4143,
            xlIBeam = 3,
            xlNorthwestArrow = 1,
            xlWait = 2,
        }
        #endregion

        #region XlMSApplication
        public enum XlMSApplication
        {
            xlMicrosoftAccess = 4,
            xlMicrosoftFoxPro = 5,
            xlMicrosoftMail = 3,
            xlMicrosoftPowerPoint = 2,
            xlMicrosoftProject = 6,
            xlMicrosoftSchedulePlus = 7,
            xlMicrosoftWord = 1,
        }
        #endregion

        #region XlOartHorizontalOverflow
        public enum XlOartHorizontalOverflow
        {
            xlOartHorizontalOverflowClip = 1,
            xlOartHorizontalOverflowOverflow = 0,
        }
        #endregion

        #region XlOartVerticalOverflow
        public enum XlOartVerticalOverflow
        {
            xlOartVerticalOverflowClip = 1,
            xlOartVerticalOverflowEllipsis = 2,
            xlOartVerticalOverflowOverflow = 0,
        }
        #endregion

        #region XlObjectSize
        public enum XlObjectSize
        {
            xlFitToPage = 2,
            xlFullPage = 3,
            xlScreenSize = 1,
        }
        #endregion

        #region XlOLEType
        public enum XlOLEType
        {
            xlOLEControl = 2,
            xlOLEEmbed = 1,
            xlOLELink = 0,
        }
        #endregion

        #region XlOLEVerb
        public enum XlOLEVerb
        {
            xlVerbOpen = 2,
            xlVerbPrimary = 1,
        }
        #endregion

        #region XlOrder
        public enum XlOrder
        {
            xlDownThenOver = 1,
            xlOverThenDown = 2,
        }
        #endregion

        #region XlOrientation
        public enum XlOrientation
        {
            xlDownward = -4170,
            xlHorizontal = -4128,
            xlUpward = -4171,
            xlVertical = -4166,
        }
        #endregion

        #region XlPageBreak
        public enum XlPageBreak
        {
            xlPageBreakAutomatic = -4105,
            xlPageBreakManual = -4135,
            xlPageBreakNone = -4142,
        }
        #endregion

        #region XlPageBreakExtent
        public enum XlPageBreakExtent
        {
            xlPageBreakFull = 1,
            xlPageBreakPartial = 2,
        }
        #endregion

        #region XlPageOrientation
        public enum XlPageOrientation
        {
            xlLandscape = 2,
            xlPortrait = 1,
        }
        #endregion

        #region XlPaperSize
        public enum XlPaperSize
        {
            xlPaper10x14 = 16,
            xlPaper11x17 = 17,
            xlPaperA3 = 8,
            xlPaperA4 = 9,
            xlPaperA4Small = 10,
            xlPaperA5 = 11,
            xlPaperB4 = 12,
            xlPaperB5 = 13,
            xlPaperCsheet = 24,
            xlPaperDsheet = 25,
            xlPaperEnvelope10 = 20,
            xlPaperEnvelope11 = 21,
            xlPaperEnvelope12 = 22,
            xlPaperEnvelope14 = 23,
            xlPaperEnvelope9 = 19,
            xlPaperEnvelopeB4 = 33,
            xlPaperEnvelopeB5 = 34,
            xlPaperEnvelopeB6 = 35,
            xlPaperEnvelopeC3 = 29,
            xlPaperEnvelopeC4 = 30,
            xlPaperEnvelopeC5 = 28,
            xlPaperEnvelopeC6 = 31,
            xlPaperEnvelopeC65 = 32,
            xlPaperEnvelopeDL = 27,
            xlPaperEnvelopeItaly = 36,
            xlPaperEnvelopeMonarch = 37,
            xlPaperEnvelopePersonal = 38,
            xlPaperEsheet = 26,
            xlPaperExecutive = 7,
            xlPaperFanfoldLegalGerman = 41,
            xlPaperFanfoldStdGerman = 40,
            xlPaperFanfoldUS = 39,
            xlPaperFolio = 14,
            xlPaperLedger = 4,
            xlPaperLegal = 5,
            xlPaperLetter = 1,
            xlPaperLetterSmall = 2,
            xlPaperNote = 18,
            xlPaperQuarto = 15,
            xlPaperStatement = 6,
            xlPaperTabloid = 3,
            xlPaperUser = 256,
        }
        #endregion

        #region XlParameterDataType
        public enum XlParameterDataType
        {
            xlParamTypeBigInt = -5,
            xlParamTypeBinary = -2,
            xlParamTypeBit = -7,
            xlParamTypeChar = 1,
            xlParamTypeDate = 9,
            xlParamTypeDecimal = 3,
            xlParamTypeDouble = 8,
            xlParamTypeFloat = 6,
            xlParamTypeInteger = 4,
            xlParamTypeLongVarBinary = -4,
            xlParamTypeLongVarChar = -1,
            xlParamTypeNumeric = 2,
            xlParamTypeReal = 7,
            xlParamTypeSmallInt = 5,
            xlParamTypeTime = 10,
            xlParamTypeTimestamp = 11,
            xlParamTypeTinyInt = -6,
            xlParamTypeUnknown = 0,
            xlParamTypeVarBinary = -3,
            xlParamTypeVarChar = 12,
            xlParamTypeWChar = -8,
        }
        #endregion

        #region XlParameterType
        public enum XlParameterType
        {
            xlConstant = 1,
            xlPrompt = 0,
            xlRange = 2,
        }
        #endregion

        #region XlPasteSpecialOperation
        public enum XlPasteSpecialOperation
        {
            xlPasteSpecialOperationAdd = 2,
            xlPasteSpecialOperationDivide = 5,
            xlPasteSpecialOperationMultiply = 4,
            xlPasteSpecialOperationNone = -4142,
            xlPasteSpecialOperationSubtract = 3,
        }
        #endregion

        #region XlPasteType
        public enum XlPasteType
        {
            xlPasteAll = -4104,
            xlPasteAllExceptBorders = 7,
            xlPasteColumnWidths = 8,
            xlPasteComments = -4144,
            xlPasteFormats = -4122,
            xlPasteFormulas = -4123,
            xlPasteFormulasAndNumberFormats = 11,
            xlPasteValidation = 6,
            xlPasteValues = -4163,
            xlPasteValuesAndNumberFormats = 12,
        }
        #endregion

        #region XlPattern
        public enum XlPattern
        {
            xlPatternAutomatic = -4105,
            xlPatternChecker = 9,
            xlPatternCrissCross = 16,
            xlPatternDown = -4121,
            xlPatternGray16 = 17,
            xlPatternGray25 = -4124,
            xlPatternGray50 = -4125,
            xlPatternGray75 = -4126,
            xlPatternGray8 = 18,
            xlPatternGrid = 15,
            xlPatternHorizontal = -4128,
            xlPatternLightDown = 13,
            xlPatternLightHorizontal = 11,
            xlPatternLightUp = 14,
            xlPatternLightVertical = 12,
            xlPatternNone = -4142,
            xlPatternSemiGray75 = 10,
            xlPatternSolid = 1,
            xlPatternUp = -4162,
            xlPatternVertical = -4166,
            xlSolid = 1,
        }
        #endregion

        #region XlPhoneticAlignment
        public enum XlPhoneticAlignment
        {
            xlPhoneticAlignCenter = 2,
            xlPhoneticAlignDistributed = 3,
            xlPhoneticAlignLeft = 1,
            xlPhoneticAlignNoControl = 0,
        }
        #endregion

        #region XlPhoneticCharacterType
        public enum XlPhoneticCharacterType
        {
            xlHiragana = 2,
            xlKatakana = 1,
            xlKatakanaHalf = 0,
            xlNoConversion = 3,
        }
        #endregion

        #region XlPictureAppearance
        public enum XlPictureAppearance
        {
            xlPrinter = 2,
            xlScreen = 1,
        }
        #endregion

        #region XlPictureConvertorType
        public enum XlPictureConvertorType
        {
            xlBMP = 1,
            xlCGM = 7,
            xlDRW = 4,
            xlDXF = 5,
            xlEPS = 8,
            xlHGL = 6,
            xlPCT = 13,
            xlPCX = 10,
            xlPIC = 11,
            xlPLT = 12,
            xlTIF = 9,
            xlWMF = 2,
            xlWPG = 3,
        }
        #endregion

        #region XlPieSliceIndex
        public enum XlPieSliceIndex
        {
            xlCenterPoint = 5,
            xlInnerCenterPoint = 8,
            xlInnerClockwisePoint = 7,
            xlInnerCounterClockwisePoint = 9,
            xlMidClockwiseRadiusPoint = 4,
            xlMidCounterClockwiseRadiusPoint = 6,
            xlOuterCenterPoint = 2,
            xlOuterClockwisePoint = 3,
            xlOuterCounterClockwisePoint = 1,
        }
        #endregion

        #region XlPieSliceLocation
        public enum XlPieSliceLocation
        {
            xlHorizontalCoordinate = 1,
            xlVerticalCoordinate = 2,
        }
        #endregion

        #region XlPivotCellType
        public enum XlPivotCellType
        {
            xlPivotCellBlankCell = 9,
            xlPivotCellCustomSubtotal = 7,
            xlPivotCellDataField = 4,
            xlPivotCellDataPivotField = 8,
            xlPivotCellGrandTotal = 3,
            xlPivotCellPageFieldItem = 6,
            xlPivotCellPivotField = 5,
            xlPivotCellPivotItem = 1,
            xlPivotCellSubtotal = 2,
            xlPivotCellValue = 0,
        }
        #endregion

        #region XlPivotConditionScope
        public enum XlPivotConditionScope
        {
            xlDataFieldScope = 2,
            xlFieldsScope = 1,
            xlSelectionScope = 0,
        }
        #endregion

        #region XlPivotFieldCalculation
        public enum XlPivotFieldCalculation
        {
            xlDifferenceFrom = 2,
            xlIndex = 9,
            xlNoAdditionalCalculation = -4143,
            xlPercentDifferenceFrom = 4,
            xlPercentOf = 3,
            xlPercentOfColumn = 7,
            xlPercentOfParent = 12,
            xlPercentOfParentColumn = 11,
            xlPercentOfParentRow = 10,
            xlPercentOfRow = 6,
            xlPercentOfTotal = 8,
            xlPercentRunningTotal = 13,
            xlRankAscending = 14,
            xlRankDecending = 15,
            xlRunningTotal = 5,
        }
        #endregion

        #region XlPivotFieldDataType
        public enum XlPivotFieldDataType
        {
            xlDate = 2,
            xlNumber = -4145,
            xlText = -4158,
        }
        #endregion

        #region XlPivotFieldOrientation
        public enum XlPivotFieldOrientation
        {
            xlColumnField = 2,
            xlDataField = 4,
            xlHidden = 0,
            xlPageField = 3,
            xlRowField = 1,
        }
        #endregion

        #region XlPivotFieldRepeatLabels
        public enum XlPivotFieldRepeatLabels
        {
            xlDoNotRepeatLabels = 1,
            xlRepeatLabels = 1,
        }
        #endregion

        #region XlPivotFilterType
        public enum XlPivotFilterType
        {
            xlBefore = 31,
            xlBeforeOrEqualTo = 32,
            xlAfter = 33,
            xlAfterOrEqualTo = 34,
            xlAllDatesInPeriodJanuary = 53,
            xlAllDatesInPeriodFebruary = 54,
            xlAllDatesInPeriodMarch = 55,
            xlAllDatesInPeriodApril = 56,
            xlAllDatesInPeriodMay = 57,
            xlAllDatesInPeriodJune = 58,
            xlAllDatesInPeriodJuly = 59,
            xlAllDatesInPeriodAugust = 60,
            xlAllDatesInPeriodSeptember = 61,
            xlAllDatesInPeriodOctober = 62,
            xlAllDatesInPeriodNovember = 63,
            xlAllDatesInPeriodDecember = 64,
            xlAllDatesInPeriodQuarter1 = 49,
            xlAllDatesInPeriodQuarter2 = 50,
            xlAllDatesInPeriodQuarter3 = 51,
            xlAllDatesInPeriodQuarter4 = 52,
            xlBottomCount = 2,
            xlBottomPercent = 4,
            xlBottomSum = 6,
            xlCaptionBeginsWith = 17,
            xlCaptionContains = 21,
            xlCaptionDoesNotBeginWith = 18,
            xlCaptionDoesNotContain = 22,
            xlCaptionDoesNotEndWith = 20,
            xlCaptionDoesNotEqual = 16,
            xlCaptionEndsWith = 19,
            xlCaptionEquals = 15,
            xlCaptionIsBetween = 27,
            xlCaptionIsGreaterThan = 23,
            xlCaptionIsGreaterThanOrEqualTo = 24,
            xlCaptionIsLessThan = 25,
            xlCaptionIsLessThanOrEqualTo = 26,
            xlCaptionIsNotBetween = 28,
            xlDateBetween = 32,
            xlDateLastMonth = 41,
            xlDateLastQuarter = 44,
            xlDateLastWeek = 38,
            xlDateLastYear = 47,
            xlDateNextMonth = 39,
            xlDateNextQuarter = 42,
            xlDateNextWeek = 36,
            xlDateNextYear = 45,
            xlDateThisMonth = 40,
            xlDateThisQuarter = 43,
            xlDateThisWeek = 37,
            xlDateThisYear = 46,
            xlDateToday = 34,
            xlDateTomorrow = 33,
            xlDateYesterday = 35,
            xlNotSpecificDate = 30,
            xlSpecificDate = 29,
            xlTopCount = 1,
            xlTopPercent = 3,
            xlTopSum = 5,
            xlValueDoesNotEqual = 8,
            xlValueEquals = 7,
            xlValueIsBetween = 13,
            xlValueIsGreaterThan = 9,
            xlValueIsGreaterThanOrEqualTo = 10,
            xlValueIsLessThan = 11,
            xlValueIsLessThanOrEqualTo = 12,
            xlValueIsNotBetween = 14,
            xlYearToDate = 48,
        }
        #endregion

        #region XlPivotFormatType
        public enum XlPivotFormatType
        {
            xlPTClassic = 20,
            xlPTNone = 21,
            xlReport1 = 0,
            xlReport10 = 9,
            xlReport2 = 1,
            xlReport3 = 2,
            xlReport4 = 3,
            xlReport5 = 4,
            xlReport6 = 5,
            xlReport7 = 6,
            xlReport8 = 7,
            xlReport9 = 8,
            xlTable1 = 10,
            xlTable10 = 19,
            xlTable2 = 11,
            xlTable3 = 12,
            xlTable4 = 13,
            xlTable5 = 14,
            xlTable6 = 15,
            xlTable7 = 16,
            xlTable8 = 17,
            xlTable9 = 18,
        }
        #endregion

        #region XlPivotLineType
        public enum XlPivotLineType
        {
            xlPivotLineBlank = 3,
            xlPivotLineGrandTotal = 2,
            xlPivotLineRegular = 0,
            xlPivotLineSubtotal = 1,
        }
        #endregion

        #region XlPivotTableMissingItems
        public enum XlPivotTableMissingItems
        {
            xlMissingItemsDefault = -1,
            xlMissingItemsMax = 32500,
            xlMissingItemsMax2 = 1048576,
            xlMissingItemsNone = 0,
        }
        #endregion

        #region XlPivotTableSourceType
        public enum XlPivotTableSourceType
        {
            xlConsolidation = 3,
            xlDatabase = 1,
            xlExternal = 2,
            xlPivotTable = -4148,
            xlScenario = 4,
        }
        #endregion

        #region XlPivotTableVersionList
        public enum XlPivotTableVersionList
        {
            xlPivotTableVersion2000 = 0,
            xlPivotTableVersion10 = 1,
            xlPivotTableVersion11 = 2,
            xlPivotTableVersion12 = 3,
            xlPivotTableVersion14 = 4,
            xlPivotTableVersion15 = 5,
            xlPivotTableVersionCurrent = -1,
        }
        #endregion

        #region XlPlacement
        public enum XlPlacement
        {
            xlFreeFloating = 3,
            xlMove = 2,
            xlMoveAndSize = 1,
        }
        #endregion

        #region XlPlatform
        public enum XlPlatform
        {
            xlMacintosh = 1,
            xlMSDOS = 3,
            xlWindows = 2,
        }
        #endregion

        #region XlPortugueseReform
        public enum XlPortugueseReform
        {
            xlPortugueseBoth = 3,
            xlPortuguesePostReform = 2,
            xlPortuguesePreReform = 1,
        }
        #endregion

        #region XlPrintErrors
        public enum XlPrintErrors
        {
            xlPrintErrorsBlank = 1,
            xlPrintErrorsDash = 2,
            xlPrintErrorsDisplayed = 0,
            xlPrintErrorsNA = 3,
        }
        #endregion

        #region XlPrintLocation
        public enum XlPrintLocation
        {
            xlPrintInPlace = 16,
            xlPrintNoComments = -4142,
            xlPrintSheetEnd = 1,
        }
        #endregion

        #region XlPriority
        public enum XlPriority
        {
            xlPriorityHigh = -4127,
            xlPriorityLow = -4134,
            xlPriorityNormal = -4143,
        }
        #endregion

        #region XlPropertyDisplayedIn
        public enum XlPropertyDisplayedIn
        {
            xlDisplayPropertyInPivotTable = 1,
            xlDisplayPropertyInPivotTableAndTooltip = 3,
            xlDisplayPropertyInTooltip = 2,
        }
        #endregion

        #region XlProtectedViewCloseReason
        public enum XlProtectedViewCloseReason
        {
            xlProtectedViewCloseEdit = 1,
            xlProtectedViewCloseForced = 2,
            xlProtectedViewCloseNormal = 0,
        }
        #endregion

        #region XlProtectedViewWindowState
        public enum XlProtectedViewWindowState
        {
            xlProtectedViewWindowMaximized = 2,
            xlProtectedViewWindowMinimized = 1,
            xlProtectedViewWindowNormal = 0,
        }
        #endregion

        #region XlPTSelectionMode
        public enum XlPTSelectionMode
        {
            xlBlanks = 4,
            xlButton = 15,
            xlDataAndLabel = 0,
            xlDataOnly = 2,
            xlFirstRow = 256,
            xlLabelOnly = 1,
            xlOrigin = 3,
        }
        #endregion

        #region XlQueryType
        public enum XlQueryType
        {
            xlADORecordset = 7,
            xlDAORecordset = 2,
            xlODBCQuery = 1,
            xlOLEDBQuery = 5,
            xlTextImport = 6,
            xlWebQuery = 4,
        }
        #endregion

        #region XlQuickAnalysisMode
        public enum XlQuickAnalysisMode
        {
            xlLensOnly = 0,
            xlFormatConditions = 1,
            xlRecommendedCharts = 2,
            xlTotals = 3,
            xlTables = 4,
            xlSparklines = 5,
        }
        #endregion

        #region XlRangeAutoFormat
        public enum XlRangeAutoFormat
        {
            xlRangeAutoFormat3DEffects1 = 13,
            xlRangeAutoFormat3DEffects2 = 14,
            xlRangeAutoFormatAccounting1 = 4,
            xlRangeAutoFormatAccounting2 = 5,
            xlRangeAutoFormatAccounting3 = 6,
            xlRangeAutoFormatAccounting4 = 17,
            xlRangeAutoFormatClassic1 = 1,
            xlRangeAutoFormatClassic2 = 2,
            xlRangeAutoFormatClassic3 = 3,
            xlRangeAutoFormatClassicPivotTable = 31,
            xlRangeAutoFormatColor1 = 7,
            xlRangeAutoFormatColor2 = 8,
            xlRangeAutoFormatColor3 = 9,
            xlRangeAutoFormatList1 = 10,
            xlRangeAutoFormatList2 = 11,
            xlRangeAutoFormatList3 = 12,
            xlRangeAutoFormatLocalFormat1 = 15,
            xlRangeAutoFormatLocalFormat2 = 16,
            xlRangeAutoFormatLocalFormat3 = 19,
            xlRangeAutoFormatLocalFormat4 = 20,
            xlRangeAutoFormatNone = -4142,
            xlRangeAutoFormatPTNone = 42,
            xlRangeAutoFormatReport1 = 21,
            xlRangeAutoFormatReport10 = 30,
            xlRangeAutoFormatReport2 = 22,
            xlRangeAutoFormatReport3 = 23,
            xlRangeAutoFormatReport4 = 24,
            xlRangeAutoFormatReport5 = 25,
            xlRangeAutoFormatReport6 = 26,
            xlRangeAutoFormatReport7 = 27,
            xlRangeAutoFormatReport8 = 28,
            xlRangeAutoFormatReport9 = 29,
            xlRangeAutoFormatSimple = -4154,
            xlRangeAutoFormatTable1 = 32,
            xlRangeAutoFormatTable10 = 41,
            xlRangeAutoFormatTable2 = 33,
            xlRangeAutoFormatTable3 = 34,
            xlRangeAutoFormatTable4 = 35,
            xlRangeAutoFormatTable5 = 36,
            xlRangeAutoFormatTable6 = 37,
            xlRangeAutoFormatTable7 = 38,
            xlRangeAutoFormatTable8 = 39,
            xlRangeAutoFormatTable9 = 40,
        }
        #endregion

        #region XlRangeValueDataType
        public enum XlRangeValueDataType
        {
            xlRangeValueDefault = 10,
            xlRangeValueMSPersistXML = 12,
            xlRangeValueXMLSpreadsheet = 11,
        }
        #endregion

        #region XlReferenceStyle
        public enum XlReferenceStyle
        {
            xlA1 = 1,
            xlR1C1 = -4150,
        }
        #endregion

        #region XlReferenceType
        public enum XlReferenceType
        {
            xlAbsolute = 1,
            xlAbsRowRelColumn = 2,
            xlRelative = 4,
            xlRelRowAbsColumn = 3,
        }
        #endregion

        #region XlRemoveDocInfoType
        public enum XlRemoveDocInfoType
        {
            xlRDIAll = 99,
            xlRDIComments = 1,
            xlRDIContentType = 16,
            xlRDIDefinedNameComments = 18,
            xlRDIDocumentManagementPolicy = 15,
            xlRDIDocumentProperties = 8,
            xlRDIDocumentServerProperties = 14,
            xlRDIDocumentWorkspace = 10,
            xlRDIEmailHeader = 5,
            xlRDIExcelDataModel = 23,
            xlRDIInactiveDataConnections = 19,
            xlRDIInkAnnotations = 11,
            xlRDIInlineWebExtensions = 21,
            xlRDIPrinterPath = 20,
            xlRDIPublishInfo = 13,
            xlRDIRemovePersonalInformation = 4,
            xlRDIRoutingSlip = 6,
            xlRDIScenarioComments = 12,
            xlRDISendForReview = 7,
            xlRDITaskpaneWebExtensions = 22,
        }
        #endregion

        #region XlRgbColor
        public enum XlRgbColor
        {
            rgbAliceBlue = 16775408,
            rgbAntiqueWhite = 14150650,
            rgbAqua = 16776960,
            rgbAquamarine = 13959039,
            rgbAzure = 16777200,
            rgbBeige = 14480885,
            rgbBisque = 12903679,
            rgbBlack = 0,
            rgbBlanchedAlmond = 13495295,
            rgbBlue = 16711680,
            rgbBlueViolet = 14822282,
            rgbBrown = 2763429,
            rgbBurlyWood = 8894686,
            rgbCadetBlue = 10526303,
            rgbChartreuse = 65407,
            rgbCoral = 5275647,
            rgbCornflowerBlue = 15570276,
            rgbCornsilk = 14481663,
            rgbCrimson = 3937500,
            rgbDarkBlue = 9109504,
            rgbDarkCyan = 9145088,
            rgbDarkGoldenrod = 755384,
            rgbDarkGray = 11119017,
            rgbDarkGreen = 25600,
            rgbDarkGrey = 11119017,
            rgbDarkKhaki = 7059389,
            rgbDarkMagenta = 9109643,
            rgbDarkOliveGreen = 3107669,
            rgbDarkOrange = 36095,
            rgbDarkOrchid = 13382297,
            rgbDarkRed = 139,
            rgbDarkSalmon = 8034025,
            rgbDarkSeaGreen = 9419919,
            rgbDarkSlateBlue = 9125192,
            rgbDarkSlateGray = 5197615,
            rgbDarkSlateGrey = 5197615,
            rgbDarkTurquoise = 13749760,
            rgbDarkViolet = 13828244,
            rgbDeepPink = 9639167,
            rgbDeepSkyBlue = 16760576,
            rgbDimGray = 6908265,
            rgbDimGrey = 6908265,
            rgbDodgerBlue = 16748574,
            rgbFireBrick = 2237106,
            rgbFloralWhite = 15792895,
            rgbForestGreen = 2263842,
            rgbFuchsia = 16711935,
            rgbGainsboro = 14474460,
            rgbGhostWhite = 16775416,
            rgbGold = 55295,
            rgbGoldenrod = 2139610,
            rgbGray = 8421504,
            rgbGreen = 32768,
            rgbGreenYellow = 3145645,
            rgbGrey = 8421504,
            rgbHoneydew = 15794160,
            rgbHotPink = 11823615,
            rgbIndianRed = 6053069,
            rgbIndigo = 8519755,
            rgbIvory = 15794175,
            rgbKhaki = 9234160,
            rgbLavender = 16443110,
            rgbLavenderBlush = 16118015,
            rgbLawnGreen = 64636,
            rgbLemonChiffon = 13499135,
            rgbLightBlue = 15128749,
            rgbLightCoral = 8421616,
            rgbLightCyan = 9145088,
            rgbLightGoldenrodYellow = 13826810,
            rgbLightGray = 13882323,
            rgbLightGreen = 9498256,
            rgbLightGrey = 13882323,
            rgbLightPink = 12695295,
            rgbLightSalmon = 8036607,
            rgbLightSeaGreen = 11186720,
            rgbLightSkyBlue = 16436871,
            rgbLightSlateGray = 10061943,
            rgbLightSteelBlue = 14599344,
            rgbLightYellow = 14745599,
            rgbLime = 65280,
            rgbLimeGreen = 3329330,
            rgbLinen = 15134970,
            rgbMaroon = 128,
            rgbMediumAquamarine = 11206502,
            rgbMediumBlue = 13434880,
            rgbMediumOrchid = 13850042,
            rgbMediumPurple = 14381203,
            rgbMediumSeaGreen = 7451452,
            rgbMediumSlateBlue = 15624315,
            rgbMediumSpringGreen = 10156544,
            rgbMediumTurquoise = 13422920,
            rgbMediumVioletRed = 8721863,
            rgbMidnightBlue = 7346457,
            rgbMintCream = 16449525,
            rgbMistyRose = 14804223,
            rgbMoccasin = 11920639,
            rgbNavajoWhite = 11394815,
            rgbNavy = 8388608,
            rgbNavyBlue = 8388608,
            rgbOldLace = 15136253,
            rgbOlive = 32896,
            rgbOliveDrab = 2330219,
            rgbOrange = 42495,
            rgbOrangeRed = 17919,
            rgbOrchid = 14053594,
            rgbPaleGoldenrod = 7071982,
            rgbPaleGreen = 10025880,
            rgbPaleTurquoise = 15658671,
            rgbPaleVioletRed = 9662683,
            rgbPapayaWhip = 14020607,
            rgbPeachPuff = 12180223,
            rgbPeru = 4163021,
            rgbPink = 13353215,
            rgbPlum = 14524637,
            rgbPowderBlue = 15130800,
            rgbPurple = 8388736,
            rgbRed = 255,
            rgbRosyBrown = 9408444,
            rgbRoyalBlue = 14772545,
            rgbSalmon = 7504122,
            rgbSandyBrown = 6333684,
            rgbSeaGreen = 5737262,
            rgbSeashell = 15660543,
            rgbSienna = 2970272,
            rgbSilver = 12632256,
            rgbSkyBlue = 15453831,
            rgbSlateBlue = 13458026,
            rgbSlateGray = 9470064,
            rgbSnow = 16448255,
            rgbSpringGreen = 8388352,
            rgbSteelBlue = 11829830,
            rgbTan = 9221330,
            rgbTeal = 8421376,
            rgbThistle = 14204888,
            rgbTomato = 4678655,
            rgbTurquoise = 13688896,
            rgbViolet = 15631086,
            rgbWheat = 11788021,
            rgbWhite = 16777215,
            rgbWhiteSmoke = 16119285,
            rgbYellow = 65535,
            rgbYellowGreen = 3329434,
        }
        #endregion

        #region XlRobustConnect
        public enum XlRobustConnect
        {
            xlAlways = 1,
            xlAsRequired = 0,
            xlNever = 2,
        }
        #endregion

        #region XlRoutingSlipDelivery
        public enum XlRoutingSlipDelivery
        {
            xlAllAtOnce = 2,
            xlOneAfterAnother = 1,
        }
        #endregion

        #region XlRoutingSlipStatus
        public enum XlRoutingSlipStatus
        {
            xlNotYetRouted = 0,
            xlRoutingComplete = 2,
            xlRoutingInProgress = 1,
        }
        #endregion

        #region XlRowCol
        public enum XlRowCol
        {
            xlColumns = 2,
            xlRows = 1,
        }
        #endregion

        #region XlRunAutoMacro
        public enum XlRunAutoMacro
        {
            xlAutoActivate = 3,
            xlAutoClose = 2,
            xlAutoDeactivate = 4,
            xlAutoOpen = 1,
        }
        #endregion

        #region XlSaveAction
        public enum XlSaveAction
        {
            xlDoNotSaveChanges = 2,
            xlSaveChanges = 1,
        }
        #endregion

        #region XlSaveAsAccessMode
        public enum XlSaveAsAccessMode
        {
            xlExclusive = 3,
            xlNoChange = 1,
            xlShared = 2,
        }
        #endregion

        #region XlSaveConflictResolution
        public enum XlSaveConflictResolution
        {
            xlLocalSessionChanges = 2,
            xlOtherSessionChanges = 3,
            xlUserResolution = 1,
        }
        #endregion

        #region XlScaleType
        public enum XlScaleType
        {
            xlScaleLinear = -4132,
            xlScaleLogarithmic = -4133,
        }
        #endregion

        #region XlSearchDirection
        public enum XlSearchDirection
        {
            xlNext = 1,
            xlPrevious = 2,
        }
        #endregion

        #region XlSearchOrder
        public enum XlSearchOrder
        {
            xlByColumns = 2,
            xlByRows = 1,
        }
        #endregion

        #region XlSearchWithin
        public enum XlSearchWithin
        {
            xlWithinSheet = 1,
            xlWithinWorkbook = 2,
        }
        #endregion

        #region XlSeriesNameLevel
        public enum XlSeriesNameLevel
        {
            xlSeriesNameLevelAll = -1,
            xlSeriesNameLevelCustom = -2,
            xlSeriesNameLevelNone = -3,
        }
        #endregion

        #region XlSheetType
        public enum XlSheetType
        {
            xlChart = -4109,
            xlDialogSheet = -4116,
            xlExcel4IntlMacroSheet = 4,
            xlExcel4MacroSheet = 3,
            xlWorksheet = -4167,
        }
        #endregion

        #region XlSheetVisibility
        public enum XlSheetVisibility
        {
            xlSheetHidden = 0,
            xlSheetVeryHidden = 2,
            xlSheetVisible = -1,
        }
        #endregion

        #region XlSizeRepresents
        public enum XlSizeRepresents
        {
            xlSizeIsArea = 1,
            xlSizeIsWidth = 2,
        }
        #endregion

        #region XlSlicerCacheType
        public enum XlSlicerCacheType
        {
            xlSlicer = 1,
            xlTimeline = 2,
        }
        #endregion

        #region XlSlicerCrossFilterType
        public enum XlSlicerCrossFilterType
        {
            xlSlicerCrossFilterHideButtonsWithNoData = 4,
            xlSlicerCrossFilterShowItemsWithDataAtTop = 2,
            xlSlicerCrossFilterShowItemsWithNoData = 3,
            xlSlicerNoCrossFilter = 1,
        }
        #endregion

        #region XlSlicerSort
        public enum XlSlicerSort
        {
            xlSlicerSortAscending = 2,
            xlSlicerSortDataSourceOrder = 1,
            xlSlicerSortDescending = 3,
        }
        #endregion

        #region XlSmartTagControlType
        public enum XlSmartTagControlType
        {
            xlSmartTagControlActiveX = 13,
            xlSmartTagControlButton = 6,
            xlSmartTagControlCheckbox = 9,
            xlSmartTagControlCombo = 12,
            xlSmartTagControlHelp = 3,
            xlSmartTagControlHelpURL = 4,
            xlSmartTagControlImage = 8,
            xlSmartTagControlLabel = 7,
            xlSmartTagControlLink = 2,
            xlSmartTagControlListbox = 11,
            xlSmartTagControlRadioGroup = 14,
            xlSmartTagControlSeparator = 5,
            xlSmartTagControlSmartTag = 1,
            xlSmartTagControlTextbox = 10,
        }
        #endregion

        #region XlSmartTagDisplayMode
        public enum XlSmartTagDisplayMode
        {
            xlButtonOnly = 2,
            xlDisplayNone = 1,
            xlIndicatorAndButton = 0,
        }
        #endregion

        #region XlSortDataOption
        public enum XlSortDataOption
        {
            xlSortNormal = 0,
            xlSortTextAsNumbers = 1,
        }
        #endregion

        #region XlSortMethod
        public enum XlSortMethod
        {
            xlPinYin = 1,
            xlStroke = 2,
        }
        #endregion

        #region XlSortMethodOld
        public enum XlSortMethodOld
        {
            xlCodePage = 2,
            xlSyllabary = 1,
        }
        #endregion

        #region XlSortOn
        public enum XlSortOn
        {
            SortOnCellColor = 1,
            SortOnFontColor = 2,
            SortOnIcon = 3,
            SortOnValues = 0,
        }
        #endregion

        #region XlSortOrder
        public enum XlSortOrder
        {
            xlAscending = 1,
            xlDescending = 2,
        }
        #endregion

        #region XlSortOrientation
        public enum XlSortOrientation
        {
            xlSortColumns = 1,
            xlSortRows = 2,
        }
        #endregion

        #region XlSortType
        public enum XlSortType
        {
            xlSortLabels = 2,
            xlSortValues = 0,
        }
        #endregion

        #region XlSourceType
        public enum XlSourceType
        {
            xlSourceAutoFilter = 3,
            xlSourceChart = 5,
            xlSourcePivotTable = 6,
            xlSourcePrintArea = 2,
            xlSourceQuery = 7,
            xlSourceRange = 4,
            xlSourceSheet = 1,
            xlSourceWorkbook = 0,
        }
        #endregion

        #region XlSpanishModes
        public enum XlSpanishModes
        {
            xlSpanishTuteoAndVoseo = 1,
            xlSpanishTuteoOnly = 0,
            xlSpanishVoseoOnly = 2,
        }
        #endregion

        #region XlSparklineRowCol
        public enum XlSparklineRowCol
        {
            SparklineColumnsSquare = 2,
            SparklineNonSquare = 0,
            SparklineRowsSquare = 1,
        }
        #endregion

        #region XlSparkScale
        public enum XlSparkScale
        {
            xlSparkScaleCustom = 3,
            xlSparkScaleGroup = 1,
            xlSparkScaleSingle = 2,
        }
        #endregion

        #region XlSparkType
        public enum XlSparkType
        {
            xlSparkColumn = 2,
            xlSparkColumnStacked100 = 3,
            xlSparkLine = 1,
        }
        #endregion

        #region XlSpeakDirection
        public enum XlSpeakDirection
        {
            xlSpeakByColumns = 1,
            xlSpeakByRows = 0,
        }
        #endregion

        #region XlSpecialCellsValue
        public enum XlSpecialCellsValue
        {
            xlErrors = 16,
            xlLogical = 4,
            xlNumbers = 1,
            xlTextValues = 2,
        }
        #endregion

        #region XlStdColorScale
        public enum XlStdColorScale
        {
            ColorScaleBlackWhite = 3,
            ColorScaleGYR = 2,
            ColorScaleRYG = 1,
            ColorScaleWhiteBlack = 4,
        }
        #endregion

        #region XlSubscribeToFormat
        public enum XlSubscribeToFormat
        {
            xlSubscribeToPicture = -4147,
            xlSubscribeToText = -4158,
        }
        #endregion

        #region XlSubtototalLocationType
        public enum XlSubtototalLocationType
        {
            xlAtBottom = 2,
            xlAtTop = 1,
        }
        #endregion

        #region XlSummaryColumn
        public enum XlSummaryColumn
        {
            xlSummaryOnLeft = -4131,
            xlSummaryOnRight = -4152,
        }
        #endregion

        #region XlSummaryReportType
        public enum XlSummaryReportType
        {
            xlStandardSummary = 1,
            xlSummaryPivotTable = -4148,
        }
        #endregion

        #region XlSummaryRow
        public enum XlSummaryRow
        {
            xlSummaryAbove = 0,
            xlSummaryBelow = 1,
        }
        #endregion

        #region XlTableStyleElementType
        public enum XlTableStyleElementType
        {
            xlBlankRow = 19,
            xlColumnStripe1 = 7,
            xlColumnStripe2 = 8,
            xlColumnSubheading1 = 20,
            xlColumnSubheading2 = 21,
            xlColumnSubheading3 = 22,
            xlFirstColumn = 3,
            xlFirstHeaderCell = 9,
            xlFirstTotalCell = 11,
            xlGrandTotalColumn = 4,
            xlGrandTotalRow = 2,
            xlHeaderRow = 1,
            xlLastColumn = 4,
            xlLastHeaderCell = 10,
            xlLastTotalCell = 12,
            xlPageFieldLabels = 26,
            xlPageFieldValues = 27,
            xlRowStripe1 = 5,
            xlRowStripe2 = 6,
            xlRowSubheading1 = 23,
            xlRowSubheading2 = 24,
            xlRowSubheading3 = 25,
            xlSlicerHoveredSelectedItemWithData = 33,
            xlSlicerHoveredSelectedItemWithNoData = 35,
            xlSlicerHoveredUnselectedItemWithData = 32,
            xlSlicerHoveredUnselectedItemWithNoData = 34,
            xlSlicerSelectedItemWithData = 30,
            xlSlicerSelectedItemWithNoData = 31,
            xlSlicerUnselectedItemWithData = 28,
            xlSlicerUnselectedItemWithNoData = 29,
            xlSubtotalColumn1 = 13,
            xlSubtotalColumn2 = 14,
            xlSubtotalColumn3 = 15,
            xlSubtotalRow1 = 16,
            xlSubtotalRow2 = 17,
            xlSubtotalRow3 = 18,
            xlTimelinePeriodLabels1 = 38,
            xlTimelinePeriodLabels2 = 39,
            xlTimelineSelectedTimeBlock = 40,
            xlTimelineSelectedTimeBlockSpace = 42,
            xlTimelineSelectionLabel = 36,
            xlTimelineTimeLevel = 37,
            xlTimelineUnselectedTimeBlock = 41,
            xlTotalRow = 2,
            xlWholeTable = 0,
        }
        #endregion

        #region XlTabPosition
        public enum XlTabPosition
        {
            xlTabPositionFirst = 0,
            xlTabPositionLast = 1,
        }
        #endregion

        #region XlTextParsingType
        public enum XlTextParsingType
        {
            xlDelimited = 1,
            xlFixedWidth = 2,
        }
        #endregion

        #region XlTextQualifier
        public enum XlTextQualifier
        {
            xlTextQualifierDoubleQuote = 1,
            xlTextQualifierNone = -4142,
            xlTextQualifierSingleQuote = 2,
        }
        #endregion

        #region XlTextVisualLayoutType
        public enum XlTextVisualLayoutType
        {
            xlTextVisualLTR = 1,
            xlTextVisualRTL = 2,
        }
        #endregion

        #region XlThemeColor
        public enum XlThemeColor
        {
            xlThemeColorAccent1 = 5,
            xlThemeColorAccent2 = 6,
            xlThemeColorAccent3 = 7,
            xlThemeColorAccent4 = 8,
            xlThemeColorAccent5 = 9,
            xlThemeColorAccent6 = 10,
            xlThemeColorDark1 = 1,
            xlThemeColorDark2 = 3,
            xlThemeColorFollowedHyperlink = 12,
            xlThemeColorHyperlink = 11,
            xlThemeColorLight1 = 2,
            xlThemeColorLight2 = 4,
        }
        #endregion

        #region XlThemeFont
        public enum XlThemeFont
        {
            xlThemeFontMajor = 2,
            xlThemeFontMinor = 1,
            xlThemeFontNone = 0,
        }
        #endregion

        #region XlThreadMode
        public enum XlThreadMode
        {
            xlThreadModeAutomatic = 0,
            xlThreadModeManual = 1,
        }
        #endregion

        #region XlTickLabelOrientation
        public enum XlTickLabelOrientation
        {
            xlTickLabelOrientationAutomatic = -4105,
            xlTickLabelOrientationDownward = -4170,
            xlTickLabelOrientationHorizontal = -4128,
            xlTickLabelOrientationUpward = -4171,
            xlTickLabelOrientationVertical = -4166,
        }
        #endregion

        #region XlTickLabelPosition
        public enum XlTickLabelPosition
        {
            xlTickLabelPositionHigh = -4127,
            xlTickLabelPositionLow = -4134,
            xlTickLabelPositionNextToAxis = 4,
            xlTickLabelPositionNone = -4142,
        }
        #endregion

        #region XlTickMark
        public enum XlTickMark
        {
            xlTickMarkCross = 4,
            xlTickMarkInside = 2,
            xlTickMarkNone = -4142,
            xlTickMarkOutside = 3,
        }
        #endregion

        #region XlTimelineLevel
        public enum XlTimelineLevel
        {
            xlTimelineLevelYears = 0,
            xlTimelineLevelQuarters = 1,
            xlTimelineLevelMonths = 2,
            xlTimelineLevelDays = 3,
        }
        #endregion

        #region XlTimePeriods
        public enum XlTimePeriods
        {
            xlLast7Days = 2,
            xlLastMonth = 5,
            xlLastWeek = 4,
            xlNextMonth = 8,
            xlNextWeek = 7,
            xlThisMonth = 9,
            xlThisWeek = 3,
            xlToday = 0,
            xlTomorrow = 6,
            xlYesterday = 1,
        }
        #endregion

        #region XlTimeUnit
        public enum XlTimeUnit
        {
            xlDays = 0,
            xlMonths = 1,
            xlYears = 2,
        }
        #endregion

        #region XlToolbarProtection
        public enum XlToolbarProtection
        {
            xlNoButtonChanges = 1,
            xlNoChanges = 4,
            xlNoDockingChanges = 3,
            xlNoShapeChanges = 2,
            xlToolbarProtectionNone = -4143,
        }
        #endregion

        #region XlTopBottom
        public enum XlTopBottom
        {
            xlTop10Bottom = 0,
            xlTop10Top = 1,
        }
        #endregion

        #region XlTotalsCalculation
        public enum XlTotalsCalculation
        {
            xlTotalsCalculationAverage = 2,
            xlTotalsCalculationCount = 3,
            xlTotalsCalculationCountNums = 4,
            xlTotalsCalculationCustom = 9,
            xlTotalsCalculationMax = 6,
            xlTotalsCalculationMin = 5,
            xlTotalsCalculationNone = 0,
            xlTotalsCalculationStdDev = 7,
            xlTotalsCalculationSum = 1,
            xlTotalsCalculationVar = 8,
        }
        #endregion

        #region XlTrendlineType
        public enum XlTrendlineType
        {
            xlExponential = 5,
            xlLinear = -4132,
            xlLogarithmic = -4133,
            xlMovingAvg = 6,
            xlPolynomial = 3,
            xlPower = 4,
        }
        #endregion

        #region XlUnderlineStyle
        public enum XlUnderlineStyle
        {
            xlUnderlineStyleDouble = -4119,
            xlUnderlineStyleDoubleAccounting = 5,
            xlUnderlineStyleNone = -4142,
            xlUnderlineStyleSingle = 2,
            xlUnderlineStyleSingleAccounting = 4,
        }
        #endregion

        #region XlUpdateLinks
        public enum XlUpdateLinks
        {
            xlUpdateLinksAlways = 3,
            xlUpdateLinksNever = 2,
            xlUpdateLinksUserSetting = 1,
        }
        #endregion

        #region XlVAlign
        public enum XlVAlign
        {
            xlVAlignBottom = -4107,
            xlVAlignCenter = -4108,
            xlVAlignDistributed = -4117,
            xlVAlignJustify = -4130,
            xlVAlignTop = -4160,
        }
        #endregion

        #region XlWBATemplate
        public enum XlWBATemplate
        {
            xlWBATChart = -4109,
            xlWBATExcel4IntlMacroSheet = 4,
            xlWBATExcel4MacroSheet = 3,
            xlWBATWorksheet = -4167,
        }
        #endregion

        #region XlWebFormatting
        public enum XlWebFormatting
        {
            xlWebFormattingAll = 1,
            xlWebFormattingNone = 3,
            xlWebFormattingRTF = 2,
        }
        #endregion

        #region XlWebSelectionType
        public enum XlWebSelectionType
        {
            xlAllTables = 2,
            xlEntirePage = 1,
            xlSpecifiedTables = 3,
        }
        #endregion

        #region XlWindowState
        public enum XlWindowState
        {
            xlMaximized = -4137,
            xlMinimized = -4140,
            xlNormal = -4143,
        }
        #endregion

        #region XlWindowType
        public enum XlWindowType
        {
            xlChartAsWindow = 5,
            xlChartInPlace = 4,
            xlClipboard = 3,
            xlInfo = -4129,
            xlWorkbook = 1,
        }
        #endregion

        #region XlWindowView
        public enum XlWindowView
        {
            xlNormalView = 1,
            xlPageBreakPreview = 2,
            xlPageLayoutView = 3,
        }
        #endregion

        #region XlXLMMacroType
        public enum XlXLMMacroType
        {
            xlCommand = 2,
            xlFunction = 1,
            xlNotXLM = 3,
        }
        #endregion

        #region XlXmlExportResult
        public enum XlXmlExportResult
        {
            xlXmlExportSuccess = 0,
            xlXmlExportValidationFailed = 1,
        }
        #endregion

        #region XlXmlImportResult
        public enum XlXmlImportResult
        {
            xlXmlImportElementsTruncated = 1,
            xlXmlImportSuccess = 0,
            xlXmlImportValidationFailed = 2,
        }
        #endregion

        #region XlXmlLoadOption
        public enum XlXmlLoadOption
        {
            xlXmlLoadImportToList = 2,
            xlXmlLoadMapXml = 3,
            xlXmlLoadOpenXml = 1,
            xlXmlLoadPromptUser = 0,
        }
        #endregion

        #region XlYesNoGuess
        public enum XlYesNoGuess
        {
            xlGuess = 0,
            xlNo = 2,
            xlYes = 1,
        }
        #endregion

        #region XlModelChangeSource
        public enum XlModelChangeSource
        {
            xlChangeByExcel = 0,
            xlChangeByPowerPivotAddIn = 1,
        }
        #endregion

    }
    #endregion

    #region GlowingTextBox
    public class GlowingTextBox : Control
    {
        public TextBox TextBox;
        private GlowingTBProperties GTBP;
        private Color GlowClr;
        private bool ML = false;

        public GlowingTextBox()
        {
            TextBox = new TextBox();
            GTBP = new GlowingTBProperties();

            GTBP.GlowColor = GlowingTBProperties.ColorState.DefaultColor;
            GTBP.MarginWidth = 5;

            base.Height = 31;
            base.Width = 80;
            base.Controls.Add(TextBox);
            base.BackColor = GTBP.DefaultColor;
            base.MinimumSize = new Size(30, 30);

            GTBP.GlowSize = 15;
            GTBP.FeatherSize = 50;

            SetSizeAndLocation();

            #region EventHandlers
            this.Resize += new EventHandler(this.AdvancedTextBox_Resize);
            TextBox.GotFocus += new EventHandler(this.TextBox_GotFocus);
            TextBox.LostFocus += new EventHandler(this.TextBox_LostFocus);
            TextBox.KeyPress += new KeyPressEventHandler(this.TextBox_KeyPress);
            TextBox.KeyUp += new KeyEventHandler(this.TextBox_KeyUp);
            TextBox.TextChanged += new EventHandler(this.TextBox_TextChanged);
            this.GotFocus += new EventHandler(this.GlowingTextBox_GotFocus);
            #endregion

            Invalidate();
        }

        #region Procedures
        /// <summary>
        ///     Set the Size and Location of the TextBox according to the Panel Size.
        /// </summary>
        private void SetSizeAndLocation()
        {
            TextBox.Location = new Point(GTBP.MarginWidth, GTBP.MarginWidth);
            TextBox.Height = base.Height - GTBP.MarginWidth * 2;
            TextBox.Width = base.Width - GTBP.MarginWidth * 2;
        }

        /// <summary>
        ///     Set the Glow color of the TextBox according to selected Preset.
        /// </summary>
        /// <param name="ClrSt">
        ///     Preset Color to Show on the Glow Color.
        /// </param>
        private void SetGlowColor(GlowingTBProperties.ColorState ClrSt)
        {
            switch (ClrSt)
            {
                case GlowingTBProperties.ColorState.DefaultColor:
                    GlowClr = GTBP.DefaultColor;
                    break;
                case GlowingTBProperties.ColorState.ErrorColor:
                    GlowClr = GTBP.ErrorColor;
                    break;
                case GlowingTBProperties.ColorState.HighLightColor:
                    GlowClr = GTBP.HighlightColor;
                    break;
                case GlowingTBProperties.ColorState.Warningcolor:
                    GlowClr = GTBP.WarningColor;
                    break;
                case GlowingTBProperties.ColorState.None:
                    GlowClr = Color.Transparent;
                    break;
            }
            Invalidate();
        }

        /// <summary>
        ///     Procedure to Repaint the TextBox Glow.
        /// </summary>
        /// <param name="e">
        ///     PaintEventArgs to repaint the Control.
        /// </param>
        protected override void OnPaintBackground(PaintEventArgs e)
        {
            base.OnPaintBackground(e);
            GraphicsPath GP = new GraphicsPath();
            using (GP)
            {
                Rectangle Rect = new Rectangle(TextBox.Bounds.X, TextBox.Bounds.Y, TextBox.Bounds.Width - 1, TextBox.Bounds.Height - 1);
                Rect.Inflate(-1, -1);
                GP.AddRectangle(Rect);
                for (int i = 1; i <= GTBP.GlowSize; i += 2)
                {
                    int AGlow = Convert.ToInt32(GTBP.FeatherSize - ((GTBP.FeatherSize / GTBP.GlowSize) * i));
                    using (Pen Pen = new Pen(Color.FromArgb(AGlow, GlowClr), i) { LineJoin = LineJoin.Round })
                    {
                        e.Graphics.DrawPath(Pen, GP);
                    }
                }
            }
        }

        private void SetML()
        {
            TextBox.Multiline = this.MultiLine;
        }
        #endregion

        #region ControlEvents
        private void AdvancedTextBox_Resize(object sender, EventArgs e)
        {
            SetSizeAndLocation();
        }

        private void TextBox_GotFocus(object sender, EventArgs e)
        {
            SetGlowColor(GTBP.FocusGlowColor);
        }

        private void TextBox_LostFocus(object sender, EventArgs e)
        {
            SetGlowColor(GTBP.GlowColor);
        }

        private void TextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            EventHandler handler = KeyPressed;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        private void TextBox_KeyUp(object sender, KeyEventArgs e)
        {
            EventHandler handler = KeyReleased;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        private void TextBox_TextChanged(object sender, EventArgs e)
        {
            EventHandler handler = TextChange;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        private void GlowingTextBox_GotFocus(object sender, EventArgs e)
        {
            TextBox.Focus();
        }
        #endregion

        #region CustomEvents
        [System.ComponentModel.Description("Event fired when a key on keyboard is pressed.")]
        public event EventHandler KeyPressed;
        [System.ComponentModel.Description("Event fired when a key on keyboard is released.")]
        public event EventHandler KeyReleased;
        [System.ComponentModel.Description("Event fired when the TextBox on Control is changed.")]
        public event EventHandler TextChange;
        #endregion

        #region Properties
        /// <summary>
        ///     The Text associated with the Control.
        /// </summary>
        /// <returns>
        ///     Returns a String that represents the Text..
        /// </returns>
        [System.ComponentModel.Description("The Text associated with the Control.")]
        public override string Text
        {
            get { return TextBox.Text; }
            set
            {
                if (value != TextBox.Text)
                {
                    TextBox.Text = value;
                }
            }
        }

        /// <summary>
        ///     Defines if the TextBox is MultiLine or not.
        /// </summary>
        /// <returns>
        ///     Returns the Boolean value indicating if the MultiLine is set or not.
        /// </returns>
        [System.ComponentModel.Description("Defines if the TextBox is MultiLine or not.")]
        public bool MultiLine
        {
            get { return ML; }
            set
            {
                ML = value;
            }
        }

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        [TypeConverter(typeof(GlowingTBPropertiesConverter))]
        [System.ComponentModel.Description("Advanced properties used in the Glowing TextBox.")]
        public GlowingTBProperties GlowingTBProperties
        {
            get { return GTBP; }
            set
            {
                if (value != GTBP)
                {
                    GTBP = value;
                    SetGlowColor(GTBP.GlowColor);
                    Invalidate();
                }
            }
        }
        #endregion
    }

    #region Class to Group the properties
    public class GlowingTBProperties
    {
        private Color DColor, EColor, HColor, WColor;
        private int marg, GlSz, ftSz;
        private ColorState ColrState, FColrState;

        public GlowingTBProperties()
        {

        }

        #region Properties

        /// <summary>
        ///     The Separation within the container and the TextBox.
        /// </summary>
        /// <returns>
        ///     Returns an Integer value that represents the MarginWidth.
        /// </returns>
        [System.ComponentModel.Description("The Separation within the container and the TextBox.")]
        public int MarginWidth
        {
            get { return marg; }
            set
            {
                if (value != marg)
                {
                    marg = value;
                    //SetSizeAndLocation();
                }
            }
        }

        /// <summary>
        ///     The Color to shown on glow on Default property.
        /// </summary>
        /// <returns>
        ///     Returns the color assigned as DefaultColor.
        /// </returns>
        [System.ComponentModel.Description("The Color to shown on glow on default property.")]
        public Color DefaultColor
        {
            get { return DColor; }
            set
            {
                if (value != DColor)
                {
                    DColor = value;
                    //SetGlowColor(GlowColor);
                }
            }
        }

        /// <summary>
        ///     The Color to shown on glow on Error property.
        /// </summary>
        /// <returns>
        ///     Returns the color assigned as ErrorColor.
        /// </returns>
        [System.ComponentModel.Description("The Color to shown on glow on error property.")]
        public Color ErrorColor
        {
            get { return EColor; }
            set
            {
                if (value != EColor)
                {
                    EColor = value;
                    //SetGlowColor(GlowColor);
                }
            }
        }

        /// <summary>
        ///     The Color to show on glow on HightLight property.
        /// </summary>
        /// <returns>
        ///     Returns the color assigned as HightlightColor.
        /// </returns>
        [System.ComponentModel.Description("The Color to show on glow on highlight property.")]
        public Color HighlightColor
        {
            get { return HColor; }
            set
            {
                if (value != HColor)
                {
                    HColor = value;
                    //SetGlowColor(GlowColor);
                }
            }
        }

        /// <summary>
        ///     The Color to show on glow on Warning property.
        /// </summary>
        /// <returns>
        ///     Returns the color assigned as WarningColor.
        /// </returns>
        [System.ComponentModel.Description("The Color to show on glow on warning property.")]
        public Color WarningColor
        {
            get { return WColor; }
            set
            {
                if (value != WColor)
                {
                    WColor = value;
                    //SetGlowColor(GlowColor);
                }
            }
        }

        /// <summary>
        ///     The Preset GlowColor to show on TextBox.
        /// </summary>
        /// <returns>
        ///     Returns the preset profile as ColorState.
        /// </returns>
        [System.ComponentModel.Description("The Preset GlowColor to show on TextBox.")]
        public ColorState GlowColor
        {
            get { return ColrState; }
            set
            {
                ColrState = value;
                //SetGlowColor(ColrState);
            }
        }

        /// <summary>
        ///     The Preset GlowColor to show when TextBox has Focus.
        /// </summary>
        /// <returns>
        ///     Returns the preset profile as ColorState.
        /// </returns>
        [System.ComponentModel.Description("The Preset GlowColor to show when TextBox has Focus.")]
        public ColorState FocusGlowColor
        {
            get { return FColrState; }
            set
            {
                if (value != FColrState)
                {
                    FColrState = value;
                }
            }
        }

        /// <summary>
        ///     The Glow Size to be shown on TextBox.
        /// </summary>
        /// <returns>
        ///     Returns the value of the GlowSize as integer.
        /// </returns>
        [System.ComponentModel.Description("The Glow size to be shown on TextBox.")]
        public int GlowSize
        {
            get { return GlSz; }
            set
            {
                if (value != GlSz)
                {
                    GlSz = value;
                    //SetGlowColor(GlowColor);
                }
            }
        }

        /// <summary>
        ///     The Feather size to be shown on TextBox.
        /// </summary>
        /// <returns>
        ///     Returns the value of the FeatherSize as integer.
        /// </returns>
        [System.ComponentModel.Description("The Feather size to be shown on TextBox.")]
        public int FeatherSize
        {
            get { return ftSz; }
            set
            {
                if (value != ftSz)
                {
                    ftSz = value;
                    //SetGlowColor(GlowColor);
                }
            }
        }
        #endregion

        #region Enumerations
        public enum ColorState
        {
            None = -1,
            DefaultColor = 0,
            ErrorColor = 1,
            HighLightColor = 2,
            Warningcolor = 3
        }
        #endregion

    }
    #endregion

    #region TypeConverter
    public class GlowingTBPropertiesConverter : ExpandableObjectConverter
    {
        public override object ConvertTo(
                 ITypeDescriptorContext context,
                 CultureInfo culture,
                 object value,
                 Type destinationType)
        {
            if (destinationType == typeof(string))
            {
                return "";
            }

            return base.ConvertTo(
                context,
                culture,
                value,
                destinationType);
        }
    }
    #endregion

    #endregion

    #region MenuItem
    [System.ComponentModel.DefaultEvent("ButtonClick")]
    public class MenuItem : Control
    {
        #region Variable Declarations
        private Label MenuText;
        private PictureBox MenuPicture;

        private int prp, H, W;
        private string MText;
        private Image NImg, HImg, CImg;
        private Color NCol, HCol, CCol;
        private RotateFlipType R;
        #endregion

        public MenuItem()
        {
            MenuText = new Label();
            MenuPicture = new PictureBox();

            Proportion = 85;
            PictureWidth = 55;

            base.Height = 65;
            base.Width = 55;
            base.Controls.Add(MenuPicture);
            base.Controls.Add(MenuText);

            MenuPicture.SizeMode = PictureBoxSizeMode.StretchImage;
            MenuText.TextAlign = ContentAlignment.BottomCenter;
            MenuText.AutoSize = false;

            SetSizes();
            SetPosition();
            SetImg();
            SetColor();

            #region Handlers
            base.Resize += new EventHandler(this.MenuItemResize);
            base.MouseHover += new EventHandler(this.MenuItemHover);
            MenuPicture.MouseHover += new EventHandler(this.MenuItemHover);
            MenuPicture.MouseEnter += new EventHandler(this.MenuItemHover);
            MenuPicture.MouseMove += new MouseEventHandler(this.MenuItemHover);
            MenuText.MouseHover += new EventHandler(this.MenuItemHover);
            MenuText.MouseEnter += new EventHandler(this.MenuItemHover);
            MenuText.MouseMove += new MouseEventHandler(this.MenuItemHover);
            base.MouseDown += new MouseEventHandler(this.MenuItemClick);
            MenuPicture.MouseDown += new MouseEventHandler(this.MenuItemClick);
            MenuText.MouseDown += new MouseEventHandler(this.MenuItemClick);
            MenuPicture.MouseLeave += new EventHandler(this.MenuItemNormal);
            MenuPicture.MouseUp += new MouseEventHandler(this.MenuItemNormal);
            MenuText.MouseLeave += new EventHandler(this.MenuItemNormal);
            MenuText.MouseUp += new MouseEventHandler(this.MenuItemNormal);
            MenuPicture.Click += new EventHandler(this.MenuItemClickButton);
            MenuText.Click += new EventHandler(this.MenuItemClickButton);
            base.Click += new EventHandler(this.MenuItemClickButton);
            #endregion

        }

        #region Procedures
        /// <summary>
        ///  Set the size of the controls contained on this control.
        /// </summary>
        private void SetSizes()
        {
            MenuPicture.Height = (base.Height * Proportion / 100);
            MenuPicture.Width = PictureWidth;
            MenuText.Height = (base.Height * (100 - Proportion) / 100);
            MenuText.Width = base.Width;
        }

        /// <summary>
        /// Set the location of the components within the control.
        /// </summary>
        private void SetPosition()
        {
            MenuPicture.Location = new Point((base.Width / 2) - MenuPicture.Width / 2, 0);
            MenuText.Location = new Point(0, MenuPicture.Height);
        }

        /// <summary>
        /// Set Color of the control Label.
        /// </summary>
        private void SetColor()
        {
            base.ForeColor = TextColor;
        }

        /// <summary>
        /// Set the default Image of the control.
        /// </summary>
        private void SetImg()
        {
            MenuPicture.Image = ImageNormal;
        }

        /// <summary>
        /// Set the Text of the control.
        /// </summary>
        private void SetText()
        {
            MenuText.Text = TextMenu;
        }

        /// <summary>
        /// Sets the Picture and Rotate effect to the MenuItem.
        /// </summary>
        /// <param name="PictureMenu"></param>
        /// <param name="ColorText"></param>
        private void VisualEffect(Image PictureMenu, Color ColorText)
        {
            if (PictureMenu != null)
            {
                MenuPicture.Image = PictureMenu;
            }
            if (ColorText != null)
            {
                MenuText.ForeColor = ColorText;
            }
            RotatePicture();
        }

        /// <summary>
        ///
        /// </summary>
        private void RotatePicture()
        {
            MenuPicture.Image.RotateFlip(Rotate);
            MenuPicture.Refresh();
        }
        #endregion

        #region Properties
        [System.ComponentModel.Description("Get or sets the height percentage that the image will occupy. The remaining percentage will be occupy by the text.")]
        public int Proportion
        {
            get { return prp; }
            set
            {
                if (value != prp)
                {
                    prp = value;
                    SetSizes();
                    SetPosition();
                }
            }
        }

        [System.ComponentModel.Description("Get or sets the Text to show on the MenuItem.")]
        public string TextMenu
        {
            get { return MText; }
            set
            {
                if (value != MText)
                {
                    MText = value;
                    SetText();
                }
            }
        }

        [System.ComponentModel.Description("Get or sets the Picture of the MenuItem.")]
        public Image ImageNormal
        {
            get { return NImg; }
            set
            {
                NImg = value;
                SetImg();
            }
        }

        [System.ComponentModel.Description("Get or sets the Picture of the MenuItem when the mouse is over the control.")]
        public Image ImageOnHover
        {
            get { return HImg; }
            set
            {
                HImg = value;
            }
        }

        [System.ComponentModel.Description("Get or sets the Picture of the MenuItem when the control is clicked.")]
        public Image ImageOnClick
        {
            get { return CImg; }
            set
            {
                CImg = value;
            }
        }

        [System.ComponentModel.Description("Get or sets the Text Color of the MenuItem.")]
        public Color TextColor
        {
            get { return NCol; }
            set
            {
                NCol = value;
                SetColor();
            }
        }

        [System.ComponentModel.Description("Get or sets the Text Color of the MenuItem when the mouse is over the control.")]
        public Color TextColorOnHover
        {
            get { return HCol; }
            set
            {
                HCol = value;
            }
        }

        [System.ComponentModel.Description("Get or sets the Text Color of the MenuItem when the control is clicked.")]
        public Color TextColorOnClick
        {
            get { return CCol; }
            set
            {
                CCol = value;
            }
        }

        [System.ComponentModel.Description("Get or sets the width of the picture.")]
        public int PictureWidth
        {
            get { return W; }
            set
            {
                W = value;
                SetSizes();
                SetPosition();
            }
        }

        public RotateFlipType Rotate
        {
            get { return R; }
            set
            {
                if (value != R)
                {
                    R = value;
                    RotatePicture();
                }
            }
        }
        #endregion

        #region EventProcedures
        private void MenuItemResize(object sender, EventArgs e)
        {
            SetSizes();
            SetPosition();
        }

        private void MenuItemHover(object sender, EventArgs e)
        {
            VisualEffect(ImageOnHover, TextColorOnHover);
        }

        private void MenuItemClick(object sender, EventArgs e)
        {
            VisualEffect(ImageOnClick, TextColorOnClick);
        }

        private void MenuItemNormal(object sender, EventArgs e)
        {
            VisualEffect(ImageNormal, TextColor);
        }

        private void MenuItemClickButton(object sender, EventArgs e)
        {
            EventHandler handler = ButtonClick;
            if (handler != null)
            {
                handler(this, e);
            }
        }
        #endregion

        #region CustomEvents
        public event EventHandler ButtonClick;
        #endregion
    }
    #endregion

    #region Export to Excel
    public class Export2Excel
    {
        #region Variable Declaration
        private DataTable dt = new DataTable();
        private SourceTypex s;
        private DatagridPlus DGP;
        string pth = "*";
        #endregion

        #region Constructors
        public Export2Excel()
        {
            this.DataSource = null;
        }
        public Export2Excel(DataTable DataSource)
        {
            SetDataSource(DataSource);
        }
        public Export2Excel(DatagridPlus DataGridPlus)
        {
            SetDataSource(DataGridPlus);
        }
        public Export2Excel(DataGridView DatagridView)
        {
            SetDataSource(DatagridView);
        }
        #endregion

        #region Export Procedure
        public string Export()
        {
            string result = "";
            Thread myth;
            myth = new Thread(new System.Threading.ThreadStart(inv));
            myth.SetApartmentState(ApartmentState.STA);
            try
            {
                dynamic xlBook;
                dynamic xlSheet;
                Type xlType = Type.GetTypeFromProgID("Excel.Application");
                dynamic app = Activator.CreateInstance(xlType);
                pth = "*";

                int xlind = 1;

                xlBook = app.Workbooks.Add();
                xlSheet = xlBook.ActiveSheet;
                app.Visible = false;

                xlSheet.Range("A1").Select();

                for (int iC = 0; iC <= DataSource.Columns.Count - 1; iC++)
                {
                    xlSheet.Cells(1, iC + 1).Value = DataSource.Columns[iC].ColumnName;
                }

                for (int iX = 0; iX <= DataSource.Rows.Count - 1; iX++)
                {
                    for (int iY = 0; iY <= DataSource.Columns.Count - 1; iY++)
                    {
                        string a;
                        a = Convert.ToString(DataSource.Rows[iX][iY]);
                        if (a != null)
                        {
                            xlSheet.Cells(xlind + 1, iY + 1).value = DataSource.Rows[iX][iY].ToString();
                        }
                    }
                    xlind++;
                }

                xlSheet.Range("A1").Select();
                xlSheet.Application.Selection.CurrentRegion.Select();
                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.xlConstants.xlNone;
                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.xlConstants.xlNone;

                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;
                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlEdgeLeft).ColorIndex = 0;
                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlEdgeLeft).TintAndShade = 0;
                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlThin;

                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlEdgeTop).ColorIndex = 0;
                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlEdgeTop).TintAndShade = 0;
                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlThin;

                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlEdgeBottom).ColorIndex = 0;
                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlEdgeBottom).TintAndShade = 0;
                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlThin;

                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlEdgeRight).ColorIndex = 0;
                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlEdgeRight).TintAndShade = 0;
                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlThin;

                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlInsideVertical).ColorIndex = 0;
                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlInsideVertical).TintAndShade = 0;
                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlThin;

                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlInsideHorizontal).ColorIndex = 0;
                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlInsideHorizontal).TintAndShade = 0;
                xlSheet.Application.Selection.Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlThin;

                xlSheet.Cells.Select();
                xlSheet.Cells.EntireColumn.AutoFit();

                xlSheet.Range("A1", Convert.ToString(GetColumnName(DataSource.Columns.Count - 1)) + "1").Select();

                xlSheet.Application.Selection.Interior.Pattern = Excel.XlPattern.xlSolid;
                xlSheet.Application.Selection.Interior.PatternColorIndex = Excel.xlConstants.xlAutomatic;
                xlSheet.Application.Selection.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1;
                xlSheet.Application.Selection.Interior.TintAndShade = -0.249977111117893;
                xlSheet.Application.Selection.Interior.PatternTintAndShade = 0;

                xlSheet.Application.Selection.Font.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
                xlSheet.Application.Selection.Font.TintAndShade = 0;

                xlSheet.Application.Selection.Font.Bold = true;
                xlSheet.Range("A1").Select();

                if (SourceType == SourceTypex.DatagridPlus)
                {
                    switch (MDGP.AdvancedProperties.GroupingMode)
                    {
                        case AdvancedProperties.GroupMethod.None:
                            FormatNone(xlSheet);
                            break;
                        case AdvancedProperties.GroupMethod.Constant:
                            FormatConstant(xlSheet);
                            break;
                        case AdvancedProperties.GroupMethod.Custom:
                            FormatCustom(xlSheet);
                            break;
                    }
                }
                else
                {
                    //** ↓↓↓ Your Code goes Here ↓↓↓ **//

                }
                myth.Start();
                while (pth == "*")
                {
                    //do nothing
                }

                if (pth != "Cancel")
                {
                    xlBook.SaveAs(pth);
                    app.Visible = true;
                    app.UserControl = true;
                    ReleaseObject(app);
                    result = "The file was saved on the following path: \r" + pth;
                    pth = "*";

                }
                else
                {
                    result = "Export canceled by user.";

                    xlBook.Close(false);
                    app.Application.Quit();

                    ReleaseObject(app);
                    pth = "*";
                }
            }
            catch (Exception e)
            {
                result = "Something went wrong when exporting the document. Please contact your system administrator." + e.Message;
            }

            return result;
        }

        private void inv()
        {
            DialogResult res;
            SaveFileDialog SFD = new SaveFileDialog();
            res = SFD.ShowDialog();

            if (res == DialogResult.OK)
            {
                pth = SFD.FileName;
            }
            else
            {
                pth = "Cancel";
            }
        }
        #endregion

        #region Excel Formating for GroupingMode.None
        private void FormatNone(dynamic xlSheet)
        {
            string lastcol = GetColumnName(MDGP.ColumnCount - 1).ToString();
            bool alt = true;
            if (MDGP.AdvancedProperties.EnableAlternateColor)
            {
                for (int i = 2; i <= MDGP.RowCount + 1; i++)
                {
                    Color rowcolor;

                    if (alt)
                    {
                        rowcolor = MDGP.AdvancedProperties.AlternateRowColor1;
                        alt = false;
                    }
                    else
                    {
                        rowcolor = MDGP.AdvancedProperties.AlternateRowColor2;
                        alt = true;
                    }

                    xlSheet.Range("A" + Convert.ToString(i), lastcol + Convert.ToString(i)).Select();
                    xlSheet.Application.Selection.Interior.Pattern = Excel.XlPattern.xlSolid;
                    xlSheet.Application.Selection.Interior.PatternColorIndex = Excel.xlConstants.xlAutomatic;
                    xlSheet.Application.Selection.Interior.Color = rowcolor;
                    xlSheet.Application.Selection.Interior.TintAndShade = 0;
                    xlSheet.Application.Selection.Interior.PatternTintAndShade = 0;
                }
            }
        }
        #endregion

        #region Excel Formating for GroupingMode.Constant
        private void FormatConstant(dynamic xlSheet)
        {
            string lastcol = GetColumnName(MDGP.ColumnCount - 1).ToString();
            bool alt = true;
            int parent = 2;

            xlSheet.Application.ActiveSheet.Outline.AutomaticStyles = false;
            xlSheet.Application.ActiveSheet.Outline.SummaryRow = Excel.xlConstants.xlAbove;
            xlSheet.Application.ActiveSheet.Outline.SummaryColumn = Excel.xlConstants.xlLeft;

            for (int i = 2; i <= MDGP.RowCount + 1; i++)
            {
                Color rowcolor;
                if (i == parent)
                {
                    if (MDGP.AdvancedProperties.EnableAlternateColor)
                    {
                        if (alt)
                        {
                            rowcolor = MDGP.AdvancedProperties.AlternateRowColor1;
                            alt = false;
                        }
                        else
                        {
                            rowcolor = MDGP.AdvancedProperties.AlternateRowColor2;
                            alt = true;
                        }
                    }
                    else
                    {
                        rowcolor = MDGP.AdvancedProperties.ParentColor;
                    }

                    parent += MDGP.AdvancedProperties.AgroupationNumber;
                }
                else
                {
                    rowcolor = MDGP.AdvancedProperties.ChildColor;
                    xlSheet.Range("A" + Convert.ToString(i), lastcol + Convert.ToString(i)).Select();
                    xlSheet.Application.Selection.Rows.Group();
                    xlSheet.Application.Selection.EntireRow.Hidden = true;
                }

                xlSheet.Range("A" + Convert.ToString(i), lastcol + Convert.ToString(i)).Select();
                xlSheet.Application.Selection.Interior.Pattern = Excel.XlPattern.xlSolid;
                xlSheet.Application.Selection.Interior.PatternColorIndex = Excel.xlConstants.xlAutomatic;
                xlSheet.Application.Selection.Interior.Color = rowcolor;
                xlSheet.Application.Selection.Interior.TintAndShade = 0;
                xlSheet.Application.Selection.Interior.PatternTintAndShade = 0;

            }
        }
        #endregion

        #region Excel Formating for GroupingMode.Custom
        private void FormatCustom(dynamic xlSheet)
        {
            string lastcol = GetColumnName(MDGP.ColumnCount - 1).ToString();
            bool alt = true;
            string parent = "";
            int j = 0;
            xlSheet.Application.ActiveSheet.Outline.AutomaticStyles = false;
            xlSheet.Application.ActiveSheet.Outline.SummaryRow = Excel.xlConstants.xlAbove;
            xlSheet.Application.ActiveSheet.Outline.SummaryColumn = Excel.xlConstants.xlLeft;

            for (int i = 2; i <= MDGP.RowCount + 1; i++)
            {
                Color rowcolor;
                if (parent != Convert.ToString(MDGP.Rows[j].Cells[0].Value))
                {
                    if (MDGP.AdvancedProperties.EnableAlternateColor)
                    {
                        if (alt)
                        {
                            rowcolor = MDGP.AdvancedProperties.AlternateRowColor1;
                            alt = false;
                        }
                        else
                        {
                            rowcolor = MDGP.AdvancedProperties.AlternateRowColor2;
                            alt = true;
                        }
                    }
                    else
                    {
                        rowcolor = MDGP.AdvancedProperties.ParentColor;
                    }

                    parent = Convert.ToString(MDGP.Rows[j].Cells[0].Value);
                }
                else
                {
                    rowcolor = MDGP.AdvancedProperties.ChildColor;
                    xlSheet.Range("A" + Convert.ToString(i), lastcol + Convert.ToString(i)).Select();
                    xlSheet.Application.Selection.Rows.Group();
                    xlSheet.Application.Selection.EntireRow.Hidden = true;
                }

                xlSheet.Range("A" + Convert.ToString(i), lastcol + Convert.ToString(i)).Select();
                xlSheet.Application.Selection.Interior.Pattern = Excel.XlPattern.xlSolid;
                xlSheet.Application.Selection.Interior.PatternColorIndex = Excel.xlConstants.xlAutomatic;
                xlSheet.Application.Selection.Interior.Color = rowcolor;
                xlSheet.Application.Selection.Interior.TintAndShade = 0;
                xlSheet.Application.Selection.Interior.PatternTintAndShade = 0;
                j++;
            }
        }
        #endregion

        #region Get Excel Column Name
        /// <summary>
        ///     Get the Excel column name based on the index Provided.
        /// </summary>
        /// <param name="index">
        ///     Integer that represents the column.
        /// </param>
        /// <returns>
        ///     Returns the string containing the Excel column Name.
        /// </returns>
        /// <source>
        ///     http://stackoverflow.com/questions/10373561/convert-a-number-to-a-letter-in-c-sharp-for-use-in-microsoft-excel
        /// </source>
        static string GetColumnName(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            var value = "";

            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];

            value += letters[index % letters.Length];

            return value;
        }
        #endregion

        #region SetDataSource
        public void SetDataSource(DataTable DataSource)
        {
            this.DataSource = DataSource;
            this.SourceType = SourceTypex.DataTable;
        }
        public void SetDataSource(DatagridPlus DataGridPlus)
        {
            this.DataSource = DataGridPlusToDataTable(DataGridPlus);
            this.SourceType = SourceTypex.DatagridPlus;
            this.MDGP = DataGridPlus;
        }
        public void SetDataSource(DataGridView DatagridView)
        {
            this.DataSource = DataGridViewToDataTable(DatagridView);
            this.SourceType = SourceTypex.DatagridView;
        }
        #endregion

        #region Release
        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
        }
        #endregion

        #region DatagridPlus to DataTable
        /// <summary>
        /// function to Convert the data from a DatagridPlus to a DataTable.
        /// </summary>
        /// <param name="DatagridPlus">DatagridPlus containing the data to convert.</param>
        /// <returns>DataTable with the data contained on DatagridPlus provided.</returns>
        private DataTable DataGridPlusToDataTable(DatagridPlus DatagridPlus)
        {
            DataTable tab = new DataTable();
            foreach (DataGridViewColumn column in DatagridPlus.Columns)
            {
                tab.Columns.Add(column.HeaderText);
            }

            for (int i = 0; i <= DatagridPlus.Rows.Count - 1; i++)
            {
                tab.Rows.Add();
                for (int j = 0; j <= DatagridPlus.Columns.Count - 1; j++)
                {
                    tab.Rows[i][j] = DatagridPlus.Rows[i].Cells[j].Value;
                }
            }
            return tab;
        }
        #endregion

        #region DatagridView to Datatable
        /// <summary>
        /// function to Convert the data from a DataGridView to a DataTable.
        /// </summary>
        /// <param name="DatagridPlus">DataGridView containing the data to convert.</param>
        /// <returns>DataTable with the data contained on DataGridView provided.</returns>
        private DataTable DataGridViewToDataTable(DataGridView DatagridView)
        {
            DataTable tab = new DataTable();
            foreach (DataGridViewColumn column in DatagridView.Columns)
            {
                tab.Columns.Add(column.HeaderText);
            }

            for (int i = 0; i <= DatagridView.Rows.Count - 1; i++)
            {
                tab.Rows.Add();
                for (int j = 0; j <= DatagridView.Columns.Count - 1; j++)
                {
                    tab.Rows[i][j] = DatagridView.Rows[i].Cells[j].Value;
                }
            }
            return tab;
        }
        #endregion

        #region Properties
        public DataTable DataSource
        {
            get { return dt; }
            set
            {
                dt = value;
            }
        }

        public SourceTypex SourceType
        {
            get { return s; }
            set
            {
                s = value;
            }
        }

        private DatagridPlus MDGP
        {
            get { return DGP; }
            set
            {
                DGP = value;
            }
        }
        #endregion

        #region Enumeration
        public enum SourceTypex
        {
            DataTable = 1,
            DatagridPlus = 2,
            DatagridView = 3,
            DataGridNoe = 4,
        }
        #endregion

    }
    #endregion

    #region Loading Class
    public class Loading : Form
    {
        private WebBrowser wb;

        #region Constructor
        public Loading()
        {
            CreateHTML();
            base.BackColor = Color.RoyalBlue;
            base.FormBorderStyle = FormBorderStyle.None;
            base.StartPosition = FormStartPosition.CenterParent;
            base.Height = 306;
            base.Width = 534;

            wb = new WebBrowser();
            wb.Height = base.Height - 20;
            wb.Width = base.Width - 20;
            base.Controls.Add(wb);
            wb.Location = new Point(10, 10);
            wb.AllowNavigation = false;
            wb.AllowWebBrowserDrop = false;
            wb.IsWebBrowserContextMenuEnabled = false;
            wb.ScrollBarsEnabled = false;
            wb.WebBrowserShortcutsEnabled = false;

            string myFile = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "loading_dots.html");
            wb.Url = new Uri("file:///" + myFile);
        }
        #endregion

        #region Create HTML
        private void CreateHTML()
        {
            string path = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "loading_dots.html");
            if (!File.Exists(path))
            {
                string html = "";// = "Hello and Welcome" + Environment.NewLine;

                html += "<!DOCTYPE html>" + Environment.NewLine;
                html += "   <html>" + Environment.NewLine;
                html += "       <head>" + Environment.NewLine;
                html += "           <style>" + Environment.NewLine;
                html += "               body{" + Environment.NewLine;
                html += "                   background-color:#a9a9a9;" + Environment.NewLine;
                html += "               }" + Environment.NewLine + Environment.NewLine;

                html += "               #fountainG{" + Environment.NewLine;
                html += "                   position:relative;" + Environment.NewLine;
                html += "                   width:400px;" + Environment.NewLine;
                html += "                   height:112px;" + Environment.NewLine;
                html += "                   margin:auto;" + Environment.NewLine;
                html += "               }" + Environment.NewLine + Environment.NewLine;

                html += "               .fountainG{" + Environment.NewLine;
                html += "                   position:absolute;" + Environment.NewLine;
                html += "                   top:0;" + Environment.NewLine;
                html += "                   background-color:#00008b;" + Environment.NewLine;
                html += "                   width:50px;" + Environment.NewLine;
                html += "                   height:50px;" + Environment.NewLine;
                html += "                   animation-name:bounce_fountainG;" + Environment.NewLine;
                html += "                       -o-animation-name:bounce_fountainG;" + Environment.NewLine;
                html += "                       -ms-animation-name:bounce_fountainG;" + Environment.NewLine;
                html += "                       -webkit-animation-name:bounce_fountainG;" + Environment.NewLine;
                html += "                       -moz-animation-name:bounce_fountainG;" + Environment.NewLine;
                html += "                   animation-duration:1.5s;" + Environment.NewLine;
                html += "                       -o-animation-duration:1.5s;" + Environment.NewLine;
                html += "                       -ms-animation-duration:1.5s;" + Environment.NewLine;
                html += "                       -webkit-animation-duration:1.5s;" + Environment.NewLine;
                html += "                       -moz-animation-duration:1.5s;" + Environment.NewLine;
                html += "                   animation-iteration-count:infinite;" + Environment.NewLine;
                html += "                       -o-animation-iteration-count:infinite;" + Environment.NewLine;
                html += "                       -ms-animation-iteration-count:infinite;" + Environment.NewLine;
                html += "                       -webkit-animation-iteration-count:infinite;" + Environment.NewLine;
                html += "                       -moz-animation-iteration-count:infinite;" + Environment.NewLine;
                html += "                   animation-direction:normal;" + Environment.NewLine;
                html += "                       -o-animation-direction:normal;" + Environment.NewLine;
                html += "                       -ms-animation-direction:normal;" + Environment.NewLine;
                html += "                       -webkit-animation-direction:normal;" + Environment.NewLine;
                html += "                       -moz-animation-direction:normal;" + Environment.NewLine;
                html += "                   transform:scale(.3);" + Environment.NewLine;
                html += "                       -o-transform:scale(.3);" + Environment.NewLine;
                html += "                       -ms-transform:scale(.3);" + Environment.NewLine;
                html += "                       -webkit-transform:scale(.3);" + Environment.NewLine;
                html += "                       -moz-transform:scale(.3);" + Environment.NewLine;
                html += "                   border-radius:74px;" + Environment.NewLine;
                html += "                       -o-border-radius:74px;" + Environment.NewLine;
                html += "                       -ms-border-radius:74px;" + Environment.NewLine;
                html += "                       -webkit-border-radius:74px;" + Environment.NewLine;
                html += "                       -moz-border-radius:74px;" + Environment.NewLine;
                html += "               }" + Environment.NewLine + Environment.NewLine;

                html += "               #fountainG_1{" + Environment.NewLine;
                html += "                   left:0;" + Environment.NewLine;
                html += "                   animation-delay:0.6s;" + Environment.NewLine;
                html += "                       -o-animation-delay:0.6s;" + Environment.NewLine;
                html += "                       -ms-animation-delay:0.6s;" + Environment.NewLine;
                html += "                       -webkit-animation-delay:0.6s;" + Environment.NewLine;
                html += "                       -moz-animation-delay:0.6s;" + Environment.NewLine;
                html += "               }" + Environment.NewLine + Environment.NewLine;

                html += "               #fountainG_2{" + Environment.NewLine;
                html += "                   left: 50px;" + Environment.NewLine;
                html += "                   animation-delay:0.75s;" + Environment.NewLine;
                html += "                       -o-animation-delay:0.75s;" + Environment.NewLine;
                html += "                       -ms-animation-delay:0.75s;" + Environment.NewLine;
                html += "                       -webkit-animation-delay:0.75s;" + Environment.NewLine;
                html += "                       -moz-animation-delay:0.75s;" + Environment.NewLine;
                html += "               }" + Environment.NewLine + Environment.NewLine;

                html += "               #fountainG_3{" + Environment.NewLine;
                html += "                   left:101px;" + Environment.NewLine;
                html += "                   animation-delay:0.9s;" + Environment.NewLine;
                html += "                       -o-animation-delay:0.9s;" + Environment.NewLine;
                html += "                       -ms-animation-delay:0.9s;" + Environment.NewLine;
                html += "                       -webkit-animation-delay:0.9s;" + Environment.NewLine;
                html += "                       -moz-animation-delay:0.9s;" + Environment.NewLine;
                html += "               }" + Environment.NewLine + Environment.NewLine;

                html += "               #fountainG_4{" + Environment.NewLine;
                html += "                   left:151px;" + Environment.NewLine;
                html += "                   animation-delay:1.05s;" + Environment.NewLine;
                html += "                       -o-animation-delay:1.05s;" + Environment.NewLine;
                html += "                       -ms-animation-delay:1.05s;" + Environment.NewLine;
                html += "                       -webkit-animation-delay:1.05s;" + Environment.NewLine;
                html += "                       -moz-animation-delay:1.05s;" + Environment.NewLine;
                html += "               }" + Environment.NewLine + Environment.NewLine;

                html += "               #fountainG_5{" + Environment.NewLine;
                html += "                   left:201px;" + Environment.NewLine;
                html += "                   animation-delay:1.2s;" + Environment.NewLine;
                html += "                       -o-animation-delay:1.2s;" + Environment.NewLine;
                html += "                       -ms-animation-delay:1.2s;" + Environment.NewLine;
                html += "                       -webkit-animation-delay:1.2s;" + Environment.NewLine;
                html += "                       -moz-animation-delay:1.2s;" + Environment.NewLine;
                html += "               }" + Environment.NewLine + Environment.NewLine;

                html += "               #fountainG_6{" + Environment.NewLine;
                html += "                   left:251px;" + Environment.NewLine;
                html += "                   animation-delay:1.35s;" + Environment.NewLine;
                html += "                       -o-animation-delay:1.35s;" + Environment.NewLine;
                html += "                       -ms-animation-delay:1.35s;" + Environment.NewLine;
                html += "                       -webkit-animation-delay:1.35s;" + Environment.NewLine;
                html += "                       -moz-animation-delay:1.35s;" + Environment.NewLine;
                html += "               }" + Environment.NewLine + Environment.NewLine;

                html += "               #fountainG_7{" + Environment.NewLine;
                html += "                   left:301px;" + Environment.NewLine;
                html += "                   animation-delay:1.5s;" + Environment.NewLine;
                html += "                       -o-animation-delay:1.5s;" + Environment.NewLine;
                html += "                       -ms-animation-delay:1.5s;" + Environment.NewLine;
                html += "                       -webkit-animation-delay:1.5s;" + Environment.NewLine;
                html += "                       -moz-animation-delay:1.5s;" + Environment.NewLine;
                html += "               }" + Environment.NewLine + Environment.NewLine;

                html += "               #fountainG_8{" + Environment.NewLine;
                html += "                   left:351px;" + Environment.NewLine;
                html += "                   animation-delay:1.64s;" + Environment.NewLine;
                html += "                       -o-animation-delay:1.64s;" + Environment.NewLine;
                html += "                       -ms-animation-delay:1.64s;" + Environment.NewLine;
                html += "                       -webkit-animation-delay:1.64s;" + Environment.NewLine;
                html += "                       -moz-animation-delay:1.64s;" + Environment.NewLine;
                html += "               }" + Environment.NewLine + Environment.NewLine;

                html += "               @keyframes bounce_fountainG{" + Environment.NewLine;
                html += "                   0%{" + Environment.NewLine;
                html += "                       transform:scale(1);" + Environment.NewLine;
                html += "                       background-color:#00008b;" + Environment.NewLine;
                html += "                   }" + Environment.NewLine + Environment.NewLine;

                html += "                   100%{" + Environment.NewLine;
                html += "                       transform:scale(.3);" + Environment.NewLine;
                html += "                       background-color:#00008b;" + Environment.NewLine;
                html += "                   }" + Environment.NewLine;
                html += "               }" + Environment.NewLine + Environment.NewLine;

                html += "               @-o-keyframes bounce_fountainG{" + Environment.NewLine;
                html += "                   0%{" + Environment.NewLine;
                html += "                       -o-transform:scale(1);" + Environment.NewLine;
                html += "                       background-color:#00008b;" + Environment.NewLine;
                html += "                   }" + Environment.NewLine + Environment.NewLine;

                html += "                   100%{" + Environment.NewLine;
                html += "                       -o-transform:scale(.3);" + Environment.NewLine;
                html += "                       background-color:#00008b;" + Environment.NewLine;
                html += "                   }" + Environment.NewLine;
                html += "               }" + Environment.NewLine + Environment.NewLine;

                html += "               @-ms-keyframes bounce_fountainG{" + Environment.NewLine;
                html += "                   0%{" + Environment.NewLine;
                html += "                       -ms-transform:scale(1);" + Environment.NewLine;
                html += "                       background-color:#00008b;" + Environment.NewLine;
                html += "                   }" + Environment.NewLine + Environment.NewLine;

                html += "                   100%{" + Environment.NewLine;
                html += "                       -ms-transform:scale(.3);" + Environment.NewLine;
                html += "                       background-color:#00008b;" + Environment.NewLine;
                html += "                   }" + Environment.NewLine;
                html += "               }" + Environment.NewLine + Environment.NewLine;

                html += "               @-webkit-keyframes bounce_fountainG{" + Environment.NewLine;
                html += "                   0%{" + Environment.NewLine;
                html += "                       -webkit-transform:scale(1);" + Environment.NewLine;
                html += "                       background-color:#00008b;" + Environment.NewLine;
                html += "                   }" + Environment.NewLine + Environment.NewLine;

                html += "                   100%{" + Environment.NewLine;
                html += "                       -webkit-transform:scale(.3);" + Environment.NewLine;
                html += "                       background-color:#00008b;" + Environment.NewLine;
                html += "                   }" + Environment.NewLine;
                html += "               }" + Environment.NewLine + Environment.NewLine;

                html += "               @-moz-keyframes bounce_fountainG{" + Environment.NewLine;
                html += "                   0%{" + Environment.NewLine;
                html += "                       -moz-transform:scale(1);" + Environment.NewLine;
                html += "                       background-color:#00008b;" + Environment.NewLine;
                html += "                   }" + Environment.NewLine + Environment.NewLine;

                html += "                   100%{" + Environment.NewLine;
                html += "                       -moz-transform:scale(.3);" + Environment.NewLine;
                html += "                       background-color:#00008b;" + Environment.NewLine;
                html += "                   }" + Environment.NewLine;
                html += "               }" + Environment.NewLine;
                html += "           </style>" + Environment.NewLine;
                html += "       </head>" + Environment.NewLine;
                html += "       <body>" + Environment.NewLine;
                html += "           <br/>" + Environment.NewLine;
                html += "           <br/>" + Environment.NewLine;
                html += "           <br/>" + Environment.NewLine;
                html += "           <br/>" + Environment.NewLine;
                html += "           <br/>" + Environment.NewLine;
                html += "           <br/>" + Environment.NewLine;
                html += "           <div id='fountainG'>" + Environment.NewLine;
                html += "               <div id='fountainG_1' class='fountainG'></div>" + Environment.NewLine;
                html += "               <div id='fountainG_2' class='fountainG'></div>" + Environment.NewLine;
                html += "               <div id='fountainG_3' class='fountainG'></div>" + Environment.NewLine;
                html += "               <div id='fountainG_4' class='fountainG'></div>" + Environment.NewLine;
                html += "               <div id='fountainG_5' class='fountainG'></div>" + Environment.NewLine;
                html += "               <div id='fountainG_6' class='fountainG'></div>" + Environment.NewLine;
                html += "               <div id='fountainG_7' class='fountainG'></div>" + Environment.NewLine;
                html += "               <div id='fountainG_8' class='fountainG'></div>" + Environment.NewLine;
                html += "           </div>" + Environment.NewLine;
                html += "       </body>" + Environment.NewLine;
                html += "   </html>" + Environment.NewLine;

                File.WriteAllText(path, html);
            }
        }
        #endregion

        #region Enumeration
        public enum ReturnType
        {
            None = 0,
            String = 1,
            Object = 2,
            Int = 3,
            Bool = 4,
            DataTable = 5,
            StringVector = 6,
            ObjectVector = 7,
            IntVector = 8,
            BoolVector = 9,
            DataTableVector = 10
        }
        #endregion
    }
    #endregion
}
