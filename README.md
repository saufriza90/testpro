# testpro

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.html;
using iTextSharp.text.html.simpleparser;


namespace XXXX
{
    public class BasePage : System.Web.UI.Page
    {
        public string constr = ConfigurationManager.ConnectionStrings["ConnString"].ConnectionString;

        private string sConnection;
        private string sql;
        private string error;
        public const string USER = "USER";
        private static int a = 0;
        private int timeout = 0;
        public const int LIMIT = 10;
        public BasePage()
        {
            sConnection = "ConnString";
        }
        protected override void OnInit(EventArgs e)
        {
            object user = Session[BasePage.USER];
            if ((user != null && user is User) == false)
            {
                Session[BasePage.USER] = null;

                Response.Redirect("Login.aspx");
                //Response.Redirect("Session.aspx");
            }
            base.OnInit(e);
        }
        
        public void redirect(string aspx)
        {
            Response.Redirect(aspx);
        }
        public void redirectPost(string url)
        {
            string[] data = url.Split('?');
            if (data.Length > 0)
            {
                url = data[0];
                Dictionary<string, string> parameter = new Dictionary<string, string>();
                if (data.Length == 2)
                {
                    string[] param = data[1].Split('&');
                    for (int i = 0; i < param.Length; i++)
                    {
                        string[] paramData = param[i].Split('=');
                        parameter[paramData[0]] = paramData[1];
                    }
                }
                redirectPost(url, parameter);
            }
        }
        public void redirectPost(string url, Dictionary<string, string> parameter)
        {
            Response.Clear();
            var sb = new System.Text.StringBuilder();
            sb.Append("<html>");
            sb.AppendFormat("<body onload='document.forms[0].submit()'>");
            sb.AppendFormat("<form action='{0}' method='post'>", url);

            foreach (KeyValuePair<string, string> entry in parameter)
            {
                sb.AppendFormat("<input type='hidden' name='" + entry.Key + "' value='{0}'>", entry.Value);
            }

            sb.Append("</form>");
            sb.Append("</body>");
            sb.Append("</html>");
            Response.Write(sb.ToString());
            Response.End();
        }
        public string get(string querystring)
        {
            if (string.IsNullOrEmpty(HttpContext.Current.Request.QueryString[querystring]))
                return string.Empty;
            else
                return HttpContext.Current.Request.QueryString[querystring];
        }
        public User getUser()
        {
            object user = Session[BasePage.USER];
            return (user is User ? (User)user : null);
        }
        public string post(string querystring)
        {
            if (string.IsNullOrEmpty(HttpContext.Current.Request.Form[querystring]))
                return string.Empty;
            else
                return HttpContext.Current.Request.Form[querystring];
        }
        public string ConnectionString
        {
            set
            {
                sConnection = value;
            }
        }
        public string ErrorMsg
        {
            get
            {
                return error;
            }
        }
        public string SqlString
        {
            set
            {
                sql = value;
            }
        }
        public int TimeOut
        {
            set
            {
                timeout = value;
            }
        }

        public void PopulateCombo(DropDownList ddl, DataTable dt, string fieldID, string fieldText)
        {
            ddl.DataSource = dt;
            ddl.DataValueField = fieldID;
            ddl.DataTextField = fieldText;
            ddl.DataBind();
            ddl.Items.Insert(0, new System.Web.UI.WebControls.ListItem("-- please select --", "0"));
        }
        public void PopulateCombo(DropDownList ddl, string sql, string fieldID, string fieldText)
        {
            this.sql = sql;
            DataTable dt = getDataTable();

            ddl.DataSource = dt;
            ddl.DataValueField = fieldID;
            ddl.DataTextField = fieldText;
            ddl.DataBind();

            ddl.Items.Insert(0, new System.Web.UI.WebControls.ListItem("-- please select --", "0"));
        }
        public void PopulateComboAll(DropDownList ddl, string dt, string fieldID, string fieldText)
        {
            ddl.DataSource = dt;
            ddl.DataValueField = fieldID;
            ddl.DataTextField = fieldText;
            ddl.DataBind();
            ddl.Items.Insert(0, new System.Web.UI.WebControls.ListItem("All", "0"));
        }

        public void PopulateComboAll(DropDownList ddl, DataTable dt, string fieldID, string fieldText)
        {
            ddl.DataSource = dt;
            ddl.DataValueField = fieldID;
            ddl.DataTextField = fieldText;
            ddl.DataBind();
            ddl.Items.Insert(0, new System.Web.UI.WebControls.ListItem("All", "0"));
        }
        public SqlConnection GetSqlConnection()
        {
            string connstring = ConfigurationManager.ConnectionStrings[sConnection].ConnectionString;
            return new SqlConnection(connstring);
        }
        public DataTable getDataTable()
        {
            if (sql == null || sql == "")
            {
                return null;
            }
            else
            {
                try
                {
                    SqlConnection myConnection = GetSqlConnection();
                    SqlCommand cmd = new SqlCommand(sql, myConnection);
                    if (timeout > 0)
                    {
                        cmd.CommandTimeout = timeout;
                    }
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    myConnection.Close();
                    return dt;
                }
                catch (Exception e)
                {
                    error = e.Message;
                    return null;
                }
            }
        }
        public DataSet getDataSet()
        {
            if (sql == null || sql == "")
            {
                return null;
            }
            else
            {
                try
                {
                    SqlConnection myConnection = GetSqlConnection();
                    SqlCommand cmd = new SqlCommand(sql, myConnection);
                    if (timeout > 0)
                    {
                        cmd.CommandTimeout = timeout;
                    }
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataSet dt = new DataSet();
                    da.Fill(dt);
                    myConnection.Close();
                    return dt;
                }
                catch (Exception e)
                {
                    error = e.Message;
                    return null;
                }
            }
        }
        public DataRow ExecuteUnique()
        {
            if (sql == null || sql == "")
            {
                return null;
            }
            else
            {
                try
                {
                    DataRow row = null;
                    SqlConnection myConnection = GetSqlConnection();
                    SqlCommand cmd = new SqlCommand(sql, myConnection);
                    if (timeout > 0)
                    {
                        cmd.CommandTimeout = timeout;
                    }
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    if (dt.Rows.Count > 0)
                    {
                        row = dt.Rows[0];
                    }
                    myConnection.Close();
                    return row;
                }
                catch (Exception e)
                {
                    error = e.Message;
                    return null;
                }
            }
        }
        public bool ExecuteNonQuery()
        {
            if (sql == null || sql == "")
            {
                return false;
            }
            else
            {
                try
                {
                    SqlConnection myConnection = GetSqlConnection();
                    SqlCommand cmd = new SqlCommand(sql, myConnection);
                    myConnection.Open();
                    if (timeout > 0)
                    {
                        cmd.CommandTimeout = timeout;
                    }
                    cmd.ExecuteNonQuery();
                    myConnection.Close();
                    return true;
                }
                catch (Exception e)
                {
                    error = e.Message;
                    return false;
                }
            }
        }
        private List<string> sqlBatch = null;
        public void prepareBatchSql()
        {
            sqlBatch = new List<string>();
            sqlBatch.Clear();
        }

        public void addBatchSql(string sql)
        {
            sqlBatch.Add(sql);
        }
        public bool transactionBatch()
        {
            SqlConnection connection = GetSqlConnection();
            SqlTransaction transaction = null;
            try
            {
                if (connection != null)
                {
                    connection.Open();
                    SqlCommand command = connection.CreateCommand();
                    if (timeout > 0)
                    {
                        command.CommandTimeout = timeout;
                    }
                    transaction = connection.BeginTransaction("beginTransactionINSYSAPP");
                    command.Connection = connection;
                    command.Transaction = transaction;
                    sqlBatch.ForEach(delegate(string sql)
                    {
                        command.CommandText = sql;
                        command.ExecuteNonQuery();
                    });
                    transaction.Commit();
                    if (connection != null)
                    {
                        connection.Close();
                        connection = null;
                    }
                    return true;
                }
                return false;
            }
            catch (Exception e)
            {
                error = e.Message;
                try
                {
                    transaction.Rollback();
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                    throw ex;
                }
                if (connection != null)
                {
                    connection.Close();
                    connection = null;
                }
                error = e.Message;
                throw e;
            }
        }
        public void exportExcel(GridView gv, Panel pnl, DataTable dt, string nama, int x)
        {
            Response.ContentType = "application/vnd.ms-excel";
            Response.AddHeader("content-disposition", "attachment;filename=" + nama + ".xls");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            StringWriter sw = new StringWriter();
            HtmlTextWriter hw = new HtmlTextWriter(sw);
            gv.AllowPaging = false;
            gv.DataSource = dt;
            gv.DataBind();

            //Change the Header Row back to white color
            gv.HeaderRow.Style.Add("background-color", "#FFFFFF");

            //Apply style to Individual Cells
            for (int i = 0; i < gv.HeaderRow.Cells.Count; i++)
            {
                gv.HeaderRow.Cells[i].Style.Add("background-color", "green");
            }

            for (int i = 0; i < gv.Rows.Count; i++)
            {
                GridViewRow row = gv.Rows[i];

                //Change Color back to white
                row.BackColor = System.Drawing.Color.White;

                //Apply text style to each Row
                row.Attributes.Add("class", "textmode");

                //Apply style to Individual Cells of Alternating Row
                if (i % 2 != 0)
                {
                    for (int z = 0; z < gv.HeaderRow.Cells.Count; z++)
                    {
                        row.Cells[z].Style.Add("background-color", "#C2D69B");
                    }
                }
            }

            pnl.RenderControl(hw);
            //style to format numbers to string
            string style = @"<style> .textmode { mso-number-format:\@; } </style>";
            Response.Write(style);
            Response.Output.Write(sw.ToString());
            Response.Flush();
            Response.End();
        }

        public void exportExcel(GridView gv, Panel pnl, DataTable dt, string nama)
        {
            Response.ContentType = "application/vnd.ms-excel";
            Response.AddHeader("content-disposition", "attachment;filename=" + nama + ".xls");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            StringWriter sw = new StringWriter();
            HtmlTextWriter hw = new HtmlTextWriter(sw);
            gv.AllowPaging = false;
            gv.DataSource = dt;
            gv.DataBind();

            //Change the Header Row back to white color
            gv.HeaderRow.Style.Add("background-color", "#FFFFFF");

            //Apply style to Individual Cells
            for (int i = 0; i < gv.HeaderRow.Cells.Count; i++)
            {
                gv.HeaderRow.Cells[i].Style.Add("background-color", "green");
            }

            for (int i = 0; i < gv.Rows.Count; i++)
            {
                GridViewRow row = gv.Rows[i];

                //Change Color back to white
                row.BackColor = System.Drawing.Color.White;

                //Apply text style to each Row
                row.Attributes.Add("class", "textmode");

                //Apply style to Individual Cells of Alternating Row
                if (i % 2 != 0)
                {
                    for (int z = 0; z < gv.HeaderRow.Cells.Count; z++)
                    {
                        row.Cells[z].Style.Add("background-color", "#C2D69B");
                    }
                }
            }

            pnl.RenderControl(hw);
            //style to format numbers to string
            string style = @"<style> .textmode { mso-number-format:\@; } </style>";
            Response.Write(style);
            Response.Output.Write(sw.ToString());
            Response.Flush();
            Response.End();
        }

        public void exportpdf(GridView gv, Panel pnl, DataTable dt, string nama, int x)
        {
            Response.ContentType = "application/pdf";
            Response.AddHeader("content-disposition", "attachment;filename=" + nama + ".pdf");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            StringWriter sw = new StringWriter();
            HtmlTextWriter hw = new HtmlTextWriter(sw);
            gv.AllowPaging = false;
            gv.DataSource = dt;
            gv.DataBind();

            //Change the Header Row back to white color
            gv.HeaderRow.Style.Add("background-color", "#FFFFFF");

            //Apply style to Individual Cells
            for (int i = 0; i < x; i++)
            {
                gv.HeaderRow.Cells[i].Style.Add("background-color", "green");
            }

            for (int i = 0; i < gv.Rows.Count; i++)
            {
                GridViewRow row = gv.Rows[i];

                //Change Color back to white
                row.BackColor = System.Drawing.Color.White;

                //Apply text style to each Row
                row.Attributes.Add("class", "textmode");

                //Apply style to Individual Cells of Alternating Row
                if (i % 2 != 0)
                {
                    for (int z = 0; z < x; z++)
                    {
                        row.Cells[z].Style.Add("background-color", "#C2D69B");
                    }
                }
            }

            pnl.RenderControl(hw);
            StringReader sr = new StringReader(sw.ToString());
            Document pdfDoc = new Document(PageSize.A2, 7f, 7f, 7f, 0f);
            HTMLWorker htmlparser = new HTMLWorker(pdfDoc);
            PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
            pdfDoc.Open();
            htmlparser.Parse(sr);
            pdfDoc.Close();
            Response.Write(pdfDoc);
            Response.End();
        }
        public override void VerifyRenderingInServerForm(Control control)
        {
            /* Verifies that the control is rendered */
        }
        public void alert(string alert)
        {
            ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('" + alert + "');", true);
        }
        public void alert(string alert, string url)
        {
            ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('" + alert + "');window.location='" + url + "';", true);
        }
        public string QuotedStr(string str)
        {
            return string.Format("'{0}'", str.Replace("'", "''"));
        }
        public string QuotedTrimStr(string str)
        {
            return QuotedStr(str.Trim());
        }
        public void windowPopUp(string url)
        {
            ClientScript.RegisterStartupScript(this.GetType(), "windowPopUp", "window.open('" + url + "', '_blank', 'toolbar=no, scrollbars=yes, resizable=no, top=100, left=100, width=500, height=400');", true);
        }
        public void windowPopUpAndRedirect(string url, string urlredirect)
        {
            ClientScript.RegisterStartupScript(this.GetType(), "windowPopUp", "window.open('" + url + "', '_blank', 'toolbar=no, scrollbars=yes, resizable=no, top=100, left=100, width=500, height=400');window.location = '" + urlredirect + "';", true);
        }
        public void redirectMessage(string message)
        {
            var url = "MessageInfo.aspx";
            Response.Clear();
            var sb = new System.Text.StringBuilder();
            sb.Append("<html>");
            sb.AppendFormat("<body onload='document.forms[0].submit()'>");
            sb.AppendFormat("<form action='{0}' method='post'>", url);
            sb.AppendFormat("<input type='hidden' name='message' value='{0}'>", message);
            sb.Append("</form>");
            sb.Append("</body>");
            sb.Append("</html>");
            Response.Write(sb.ToString());
            Response.End();
        }
        public string numberFormat(string number)
        {
            return numberFormat(Convert.ToDouble(number == "" ? "0" : number));
        }
        public string numberFormat(double number)
        {
            return (number == 0 ? "0" : number.ToString("#,#;(#,#)"));
        }
        public string decimalFormat(string number)
        {
            return decimalFormat(Convert.ToDouble(number == "" ? "0" : number));
        }
        public string decimalFormat(double number)
        {
            return (number == 0 ? "0" : number.ToString("#,#.00;(#,#.00)"));
        }
        public string numberCustomeFormat(string format, double number)
        {
            return number.ToString(format);
        }
        public string emptyStringToZero(string number)
        {
            return emptyStringToZero(number == "" ? "0" : number);
        }

        public double EffectiveRate(string id)
        {
            double effRate = 0;
            this.ConnectionString = "ConnString";
            this.SqlString = "exec getDataNPV '" + id + "'";
            DataTable dt = getDataTable();
            if (dt.Rows.Count > 0)
            {
                double tempx = -(1 * (1 + ((Convert.ToDouble(dt.Rows[0]["FlatSupplier"].ToString()) / 100) * Convert.ToDouble(dt.Rows[0]["TENOR"].ToString()) / 12))) / Convert.ToDouble(dt.Rows[0]["TENOR"].ToString());

                if (dt.Rows[0]["FIRSTINSTALLMENT"].ToString() == "AD")
                {
                    if (Convert.ToDouble(dt.Rows[0]["TENOR"]) < 60)
                    {
                        effRate = Convert.ToDouble(12 * Microsoft.VisualBasic.Financial.Rate(Convert.ToDouble(dt.Rows[0]["TENOR"].ToString()), tempx, 1, 0, Microsoft.VisualBasic.DueDate.BegOfPeriod, 1) * 100);
                    }
                    else 
                    {
                        effRate = Convert.ToDouble(12 * Microsoft.VisualBasic.Financial.Rate(Convert.ToDouble(dt.Rows[0]["TENOR"]), Convert.ToDouble(dt.Rows[0]["INSTALLMENTAMOUNT"]) * 1000, (Convert.ToDouble(dt.Rows[0]["NTF"]) * -1) * 1000, 0, Microsoft.VisualBasic.DueDate.BegOfPeriod)) * 100;
                    }
                }
                else
                {
                    if (Convert.ToDouble(dt.Rows[0]["TENOR"]) < 60)
                    {
                        effRate = Convert.ToDouble(12 * Microsoft.VisualBasic.Financial.Rate(Convert.ToDouble(dt.Rows[0]["TENOR"].ToString()), tempx, 1, 0, Microsoft.VisualBasic.DueDate.EndOfPeriod, 1) * 100);
                    }
                    else
                    {
                        effRate = Convert.ToDouble(12 * Microsoft.VisualBasic.Financial.Rate(Convert.ToDouble(dt.Rows[0]["TENOR"]), Convert.ToDouble(dt.Rows[0]["INSTALLMENTAMOUNT"]) * 1000, (Convert.ToDouble(dt.Rows[0]["NTF"]) * -1) * 1000, 0, Microsoft.VisualBasic.DueDate.EndOfPeriod)) * 100;
                    }
                }
            }
            return effRate;
            //string Query = "exec SaveAllDataNPV '" + id + "','" + effRate + "'";
            //int x = CDA.UpdateData(Query);
        }
        public int setInt(string values)
        { 
            return Convert.ToInt32(values == "" ? "0" : values);
        }
        public double setDouble(string values)
        {
            return Convert.ToDouble(values == "" ? "0" : values);
        }
        public string setMoneyFormat(string values)
        {
            return String.Format("{0:#,##0.##}", numberFormat(values));
        }
        public string setMoneyFormat(double values)
        {
            return String.Format("{0:#,##0.##}", values);
        }
        public bool save_TrxProsesSequenceDetail_record(string RegistrationID, string ProsesID, string StepID, string CodeID, string StartEnd)
        {
            bool status = false;
            this.ConnectionString = "ConnString";
            string query = "exec TrxProsesSequenceDetail_record '" + RegistrationID + "','" + ProsesID + "','" + StepID + "','" + CodeID + "','" + StartEnd + "','" + this.Session["username"] + "','" + this.Session["position"] + "'";
            this.SqlString = query;
            this.timeout = 100;
            if (this.ExecuteNonQuery())
            {
                status = true;
            }
            return status;
        }
        public bool save_insertTrxProsesSequenceDetailSPVDesktop(string RegistrationID, string docType)
        {
            bool status = false;
            this.ConnectionString = "ConnString";
            this.TimeOut = 100;
            this.SqlString = "exec insertTrxProsesSequenceDetailSPVDesktop '" + RegistrationID + "','" + this.Session["username"] + "','" + docType + "'";
            if (this.ExecuteNonQuery())
            {
                status = true;
            }
            return status;
        }
        public bool save_InsertTrxProsesSequence(string RegistrationID, string ProsesID, string StepID, string StepSequence, string StartEnd)
        {
            bool status = false;
            this.ConnectionString = "ConnString";
            this.TimeOut = 100;
            this.SqlString = "exec InsertTrxProsesSequence '" + RegistrationID + "','" + ProsesID + "','" + StepID + "','" + StepSequence + "','" + this.Session["username"] + "','" + this.Session["position"] + "','" + StartEnd + "'";
            if (this.ExecuteNonQuery())
            {
                status = true;
            }
            return status;
        }
        public DataTable dtSeqLast(string RegistrationID)
        {
            DataTable dt = new DataTable();
            string query = "select top 1 * from TrxProsesSequence where RegistrationID ='" + RegistrationID + "' order by StartDate desc";
            this.SqlString = query;
            dt = this.getDataTable();
            return dt;
        }
        public bool save_TrxProsesSequenceDetail_record_V2(string RegistrationID, string ProsesID, string StepID, string CodeId, string StartEnd)
        {
            bool status = false;
            this.ConnectionString = "ConnString";
            this.TimeOut = 100;
            this.SqlString = "exec TrxProsesSequenceDetail_record_V2 '" + RegistrationID + "','" + ProsesID + "','" + StepID + "','" + CodeId + "','" + StartEnd + "','" + this.Session["username"] + "','" + this.Session["position"] + "'";
            if (this.ExecuteNonQuery())
            {
                status = true;
            }
            return status;
        }

        public string Base64Encode(string plainText)
        {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
            return System.Convert.ToBase64String(plainTextBytes);
        }

        public string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }
    }
}
public class User
{
    public string USERID { set; get; }
    public string BRANCH { set; get; }
    public string AREA { set; get; }
    public string POSITION { set; get; }
    public string EMP_NAME { set; get; }
    public int GROUPMENU { set; get; }
}
