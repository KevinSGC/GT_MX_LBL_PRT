using iTextSharp.text;
using iTextSharp.text.pdf;
using MX_LBL_PRT.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Management;
using System.Windows.Forms;
using static MX_LBL_PRT.Program;

namespace MX_LBL_PRT
{
    public partial class FrmPolarisLICPlate : Form
    {
        public FrmPolarisLICPlate()
        {
            InitializeComponent();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            this.DoPrint();
        }

        private void Wait(int seconds)
        {
            DateTime start = DateTime.Now;
            while (start.AddSeconds(seconds) >= DateTime.Now)
            {
                Application.DoEvents();
            }
        }

        private static bool CheckPrintQueue(string file)
        {
            //尋找PrintQueue有沒有檔案相同的列印工作
            string searchQuery =
                "SELECT * FROM Win32_PrintJob";
            //var printJobs =
            //         new ManagementObjectSearcher(searchQuery).Get();
            //return printJobs.Any(o => (string)o.Properties["Document"].Value == file);
            string jobName = string.Empty;
            ManagementObjectSearcher printJobs = new ManagementObjectSearcher(searchQuery);
            foreach (ManagementObject mo in printJobs.Get())
            {
                jobName = mo.Properties["Document"].Value.ToString();
            }

            if (jobName == file)
            { return true; }
            else
                return false;
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            //regenerate print id
            AppStatic.printID = Guid.NewGuid().ToString().ToUpper();

            //check cust pn
            if (cbbCustPN.Text == "")
            {
                MessageBox.Show("Please select Cust PN!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cbbShipFrom.Text == "")
            {
                MessageBox.Show("Please select Ship From!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (cbbShipTo.Text == "")
            {
                MessageBox.Show("Please select Ship To!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (tbCustPO0.Text == "")
            {
                MessageBox.Show("Please keyin Cust PO!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (tbRel.Text == "")
            {
                MessageBox.Show("Please keyin PO Revison!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (tbCustPO0.Text.Length!=6)
            {
                MessageBox.Show("PO length must be 6!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (tbQty0.Text == "")
            {
                MessageBox.Show("Please Keyin QTY!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                string ret = "0";
                string msg = "";
                SqlConnection conn = db.GetSqlConnection();
                conn.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = conn;

                SqlTransaction tran = conn.BeginTransaction();
                cmd.Transaction = tran;
                for (int i = 1; i <= nupAppBoxQty.Value; i++)
                {
                    cmd.CommandText = "MX_GEN_LBL_POLARIS";
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@PRT_ID", System.Data.SqlDbType.VarChar, 50).Value = AppStatic.printID;
                    cmd.Parameters.Add("@PART_NO", System.Data.SqlDbType.VarChar, 50).Value = cbbCustPN.Text;
                    cmd.Parameters.Add("@SHIP_FROM", System.Data.SqlDbType.VarChar, 50).Value = cbbShipFrom.Text;
                    cmd.Parameters.Add("@SHIP_TO", System.Data.SqlDbType.VarChar, 50).Value = cbbShipTo.Text;
                    cmd.Parameters.Add("@CUST_PO", System.Data.SqlDbType.VarChar, 50).Value = tbCustPO0.Text;
                    cmd.Parameters.Add("@REL", System.Data.SqlDbType.VarChar, 50).Value = tbRel.Text;
                    cmd.Parameters.Add("@QTY", System.Data.SqlDbType.Int).Value = tbQty0.Text;
                    cmd.Parameters.Add("@RET", System.Data.SqlDbType.VarChar, 50).Direction = System.Data.ParameterDirection.Output;
                    cmd.Parameters.Add("@MSG", System.Data.SqlDbType.VarChar, 50).Direction = System.Data.ParameterDirection.Output;
                    cmd.ExecuteNonQuery();
                    ret = cmd.Parameters["@RET"].Value.ToString();
                    msg = cmd.Parameters["@MSG"].Value.ToString();
                    if(ret=="0")
                    {
                        break;
                    }
                }
                if(ret=="1")
                {
                    tran.Commit();
                    MessageBox.Show("Label generate OK!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    tran.Rollback();
                    MessageBox.Show("Label generate failed!\r\n"+msg, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Label generate failed!\r\n" + ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            //refresh the jobs
            this.ShowPrintJobs("");

        }

        //private string printID = string.Empty;
        private DbHelper db = new DbHelper();
        private void FrmPolarisLICPlate_Load(object sender, EventArgs e)
        {
            this.ClearText();
            //this.Text = AppStatic.printID;
            using (SqlConnection conn = db.GetSqlConnection())
            {
                conn.Open();

                SqlCommand cmd = new SqlCommand();
                cmd.Connection = conn;

                //show cust pn
                cbbCustPN.Items.Clear();
                cmd.CommandText = "SELECT CUST_PN FROM MX_CUST_PARTS";
                SqlDataReader drCustPN = cmd.ExecuteReader();
                while (drCustPN.Read())
                {
                    cbbCustPN.Items.Add(drCustPN[0].ToString());
                }
                drCustPN.Close();

                //show shipfrom
                cmd.CommandText = "SELECT SHIP_FROM FROM MX_SHIP_FROM";
                SqlDataReader drShipFrom = cmd.ExecuteReader();
                while (drShipFrom.Read())
                {
                    cbbShipFrom.Items.Add(drShipFrom[0].ToString());
                }
                drShipFrom.Close();
                if (cbbShipFrom.Items.Count > 0) cbbShipFrom.SelectedIndex = 0;

                //show shipto
                cmd.CommandText = "SELECT SHIP_TO FROM MX_SHIP_TO";
                SqlDataReader drShipTo = cmd.ExecuteReader();
                while (drShipTo.Read())
                {
                    cbbShipTo.Items.Add(drShipTo[0].ToString());
                }
                drShipTo.Close();
                conn.Close();
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            //refresh the 
            this.ShowPrintJobs("");
        }

        void ShowPrintJobs(string PrintID)
        {
            using(SqlConnection conn = db.GetSqlConnection())
            {
                conn.Open();

                SqlCommand cmd = new SqlCommand();
                cmd.Connection = conn;
                if (PrintID == "")
                {
                    cmd.CommandText = "SELECT PRT_ID,SHIP_FROM,SHIP_TO,ASN_ID,CUST_PN,PART_DESC,REV,CUST_PO,QTY,UPL,LTC,CDATE,PDATE,PRT_CNT,SITE FROM MX_PRT_JOB WHERE PDATE IS NULL";
                }
                else
                {
                    cmd.CommandText = "SELECT PRT_ID,SHIP_FROM,SHIP_TO,ASN_ID,CUST_PN,PART_DESC,REV,CUST_PO,QTY,UPL,LTC,CDATE,PDATE,PRT_CNT,SITE FROM MX_PRT_JOB WHERE PDATE IS NULL AND PRT_ID='" + PrintID + "'";
                }
                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                bs1.DataSource = dt;
                bn1.BindingSource = bs1;
                dgv1.DataSource = bs1;
                conn.Close();
            }
        }

        private void dgv1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if(dgv1.SelectedCells!=null)
            {
                this.ClearText();
                string shipfrom= dgv1[1, e.RowIndex].Value.ToString();
                string shipto = dgv1[2, e.RowIndex].Value.ToString();
                //1.find the ship from address
                //2.find the ship to address
                //3.show the address
                using(SqlConnection conn = db.GetSqlConnection())
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;

                    //show shipfrom
                    cmd.CommandText = "SELECT ADDR1,ADDR2,ADDR3,ADDR4,ADDR5 FROM MX_SHIP_FROM WHERE SHIP_FROM='" + shipfrom + "'";
                    SqlDataReader drShipFrom = cmd.ExecuteReader();
                    if(drShipFrom.Read())
                    {
                        //tbShipTo.Lines[0] = drShipFrom[0].ToString();
                        //tbShipTo.Lines[1] = drShipFrom[1].ToString();
                        //tbShipTo.Lines[2] = drShipFrom[2].ToString();
                        //tbShipTo.Lines[3] = drShipFrom[3].ToString();
                        //tbShipTo.Lines[4] = drShipFrom[4].ToString();
                        tbShipFrom.AppendText(drShipFrom[0].ToString() + "\r\n");
                        tbShipFrom.AppendText(drShipFrom[1].ToString() + "\r\n");
                        tbShipFrom.AppendText(drShipFrom[2].ToString() + "\r\n");
                        tbShipFrom.AppendText(drShipFrom[3].ToString() + "\r\n");
                        tbShipFrom.AppendText(drShipFrom[4].ToString() + "\r\n");
                    }
                    drShipFrom.Close();

                    //show shipto
                    cmd.CommandText = "SELECT ADDR1,ADDR2,ADDR3,ADDR4,ADDR5 FROM MX_SHIP_TO WHERE SHIP_TO='" + shipto + "'";
                    SqlDataReader drShipTo = cmd.ExecuteReader();
                    if (drShipTo.Read())
                    {
                        //tbShipFrom.Lines[0] = drShipTo[0].ToString();
                        //tbShipFrom.Lines[1] = drShipTo[1].ToString();
                        //tbShipFrom.Lines[2] = drShipTo[2].ToString();
                        //tbShipFrom.Lines[3] = drShipTo[3].ToString();
                        //tbShipFrom.Lines[4] = drShipTo[4].ToString();
                        tbShipTo.AppendText(drShipTo[0].ToString() + "\r\n");
                        tbShipTo.AppendText(drShipTo[1].ToString() + "\r\n");
                        tbShipTo.AppendText(drShipTo[2].ToString() + "\r\n");
                        tbShipTo.AppendText(drShipTo[3].ToString() + "\r\n");
                        tbShipTo.AppendText(drShipTo[4].ToString() + "\r\n");
                    }
                    drShipTo.Close();
                }
                tbASN.Text = dgv1[3, e.RowIndex].Value.ToString();
                tbCustPart.Text= dgv1[4, e.RowIndex].Value.ToString();
                tbPartDesc.Text= dgv1[5, e.RowIndex].Value.ToString();
                tbREV.Text= dgv1[6, e.RowIndex].Value.ToString();
                tbCustPO.Text= dgv1[7, e.RowIndex].Value.ToString();
                tbQTY.Text= dgv1[8, e.RowIndex].Value.ToString();
                tbLICPlate.Text= dgv1[9, e.RowIndex].Value.ToString();
                tbLotTrace.Text= dgv1[10, e.RowIndex].Value.ToString();
            }
        }

        private void ClearText()
        {
            tbShipFrom.Clear();
            tbShipTo.Clear();
            tbASN.Clear();
            tbCustPart.Clear();
            tbPartDesc.Clear();
            tbREV.Clear();
            tbCustPO.Clear();
            tbQTY.Clear();
            tbLICPlate.Clear();
            tbLotTrace.Clear();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (tbLICPlate.Text == "") return;
            this.DoPrint();
            this.SetStatusUpdated(tbLICPlate.Text);
            this.ClearText();
            this.ShowPrintJobs("");
        }

        private void SetStatusUpdated(string upl)
        {
            using (SqlConnection conn = db.GetSqlConnection())
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = conn;
                cmd.CommandText = "UPDATE MX_PRT_JOB SET PDATE=GETDATE(),PRT_CNT=PRT_CNT+1 WHERE UPL='" + upl + "'";
                cmd.ExecuteNonQuery();
                conn.Close();
            }
        }

        void DoPrint()
        {
            try
            {
                btnPrint.Enabled = false;
                if (!Directory.Exists("PDF"))
                {
                    Directory.CreateDirectory("PDF");
                }
                var pdf = new iTextSharp.text.Document(new iTextSharp.text.Rectangle(468, 288f));
                //var pdfFile = Application.StartupPath + @"c:\temp\temp100.pdf";
                var fileName = DateTime.Now.ToString("yyyyMMddHHmmss");
                var pdfFile = @"PDF\" + fileName + @".pdf";

                //var ms = new MemoryStream();

                PdfWriter writer = PdfWriter.GetInstance(pdf, new FileStream(pdfFile, FileMode.Create));
                //PdfWriter pdfWriter = PdfWriter.GetInstance(pdf, ms);

                pdf.Open();

                //var qrBitmap = ZXingNetBarcodeHelper.GenerateQRCode("1234567890", 25, 25);
                //var txm = iTextSharp.text.Image.GetInstance(System.Drawing.Image.FromHbitmap(qrBitmap.GetHbitmap()), ImageFormat.Bmp);
                ////txm.ScaleAbsoluteHeight(40);
                //txm.SetAbsolutePosition(200, 200);
                //pdf.Add(txm);
                ////get the day of the year
                //int dayOfYear = DateTime.Now.DayOfYear;
                //var bitmapBarcode1 = ZXingNetBarcodeHelper.GenerateCode128(string.Format("2S{0}{1}9001",DateTime.Now.Year,DateTime.Now.DayOfYear), 25, 15);
                //var txm1 = iTextSharp.text.Image.GetInstance(System.Drawing.Image.FromHbitmap(bitmapBarcode1.GetHbitmap()), ImageFormat.Bmp);
                ////txm.ScaleAbsoluteHeight(40);
                //txm1.SetAbsolutePosition(100, 200);
                //pdf.Add(txm1);

                PdfContentByte cb = writer.DirectContent;
                cb.SetLineWidth(0.6f);
                cb.SetColorStroke(BaseColor.BLACK);

                cb.MoveTo(218, 5);
                cb.LineTo(218, 80);
                cb.Stroke();

                cb.MoveTo(5, 80);
                cb.LineTo(463, 80);
                cb.Stroke();

                cb.MoveTo(280, 80);
                cb.LineTo(280, 285);
                cb.Stroke();

                cb.MoveTo(5, 145);
                cb.LineTo(463, 145);
                cb.Stroke();

                cb.MoveTo(5, 210);
                cb.LineTo(463, 210);
                cb.Stroke();

                cb.MoveTo(280, 174);
                cb.LineTo(463, 174);
                cb.Stroke();

                cb.MoveTo(140, 210);
                cb.LineTo(140, 285);
                cb.Stroke();

                //remove the rect
                //cb.Rectangle(1, 1, pdf.PageSize.Width - 2, pdf.PageSize.Height - 2);
                //cb.Stroke();

                var ltc = tbLotTrace.Text;
                var lpu = tbLICPlate.Text;
                var custPN = tbCustPart.Text;
                var po = tbCustPO.Text;
                var rev = tbREV.Text;
                var qty = tbQTY.Text;
                var asn = tbASN.Text;
                var partDesc = tbPartDesc.Text;
                var custPO = tbCustPO.Text;

                List<string> shiptos = new List<string>();
                for (int i = 0; i < 5; i++)
                {
                    if (tbShipTo.Lines.Count() < i + 1)
                    {
                        shiptos.Add("");
                    }
                    else
                        shiptos.Add(tbShipTo.Lines[i]);
                }

                var bitmapBarcode1 = ZXingNetBarcodeHelper.GenerateCode128("1J" + lpu, 25, 25);
                var txm1 = iTextSharp.text.Image.GetInstance(System.Drawing.Image.FromHbitmap(bitmapBarcode1.GetHbitmap()), ImageFormat.Bmp);
                //txm.ScaleAbsoluteHeight(40);
                txm1.SetAbsolutePosition(15, 10);
                pdf.Add(txm1);

                var bitmapBarcode2 = ZXingNetBarcodeHelper.GenerateCode128("1T" + ltc, 25, 25);
                var txm2 = iTextSharp.text.Image.GetInstance(System.Drawing.Image.FromHbitmap(bitmapBarcode2.GetHbitmap()), ImageFormat.Bmp);
                //txm.ScaleAbsoluteHeight(40);
                txm2.SetAbsolutePosition(250, 24);
                txm2.ScaleAbsoluteWidth(200);
                pdf.Add(txm2);

                var bitmapBarcode3 = ZXingNetBarcodeHelper.GenerateCode128("K" + po, 25, 25);
                var txm3 = iTextSharp.text.Image.GetInstance(System.Drawing.Image.FromHbitmap(bitmapBarcode3.GetHbitmap()), ImageFormat.Bmp);
                //txm.ScaleAbsoluteHeight(40);
                txm3.SetAbsolutePosition(40, 90);
                //txm3.ScaleAbsoluteWidth(200);
                pdf.Add(txm3);

                var bitmapBarcode4 = ZXingNetBarcodeHelper.GenerateCode128("Q" + qty, 25, 25);
                var txm4 = iTextSharp.text.Image.GetInstance(System.Drawing.Image.FromHbitmap(bitmapBarcode4.GetHbitmap()), ImageFormat.Bmp);
                //txm4.ScaleAbsoluteHeight(40);
                txm4.SetAbsolutePosition(330, 111);
                //txm4.ScaleAbsoluteWidth(200);
                pdf.Add(txm4);

                var bitmapBarcode5 = ZXingNetBarcodeHelper.GenerateCode128("P" + custPN, 25, 25);
                var txm5 = iTextSharp.text.Image.GetInstance(System.Drawing.Image.FromHbitmap(bitmapBarcode5.GetHbitmap()), ImageFormat.Bmp);
                //txm5.ScaleAbsoluteHeight(40);
                txm5.SetAbsolutePosition(40, 150);
                //txm5.ScaleAbsoluteWidth(200);
                pdf.Add(txm5);

                var bitmapBarcode6 = ZXingNetBarcodeHelper.GenerateCode128("2S" + asn, 25, 25);
                var txm6 = iTextSharp.text.Image.GetInstance(System.Drawing.Image.FromHbitmap(bitmapBarcode6.GetHbitmap()), ImageFormat.Bmp);
                //txm6.ScaleAbsoluteHeight(40);
                txm6.SetAbsolutePosition(305, 238);
                //txm5.ScaleAbsoluteWidth(200);
                pdf.Add(txm6);

                var baseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

                cb.BeginText();
                cb.SetFontAndSize(baseFont, 16);
                cb.SetColorFill(BaseColor.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("{0}", lpu), 15, 40, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("{0}", ltc), 250, 10, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("{0}", po), 40, 120, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("{0}", qty), 355, 95, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("{0}", custPN), 40, 180, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("{0}", rev), 360, 155, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("{0}", asn), 305, 222, 0);
                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(baseFont, 8);
                cb.SetColorFill(BaseColor.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("UNIT(1J):"), 10, 60, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("LIC PLATE-"), 10, 70, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("(1T):"), 220, 40, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("CODE"), 220, 50, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("TRACE"), 220, 60, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("LOT"), 220, 70, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("CUST"), 10, 135, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("PO#"), 10, 125, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("(K):"), 10, 115, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("QTY(Q):"), 285, 90, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("CUST PART"), 10, 199, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("(P):"), 10, 189, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("REV"), 285, 163, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("LEVEL:"), 285, 153, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("PART"), 285, 199, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("DESC:"), 285, 189, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("ASN ID:"), 285, 275, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("(2S)"), 285, 267, 0);

                cb.EndText();

                cb.BeginText();
                cb.SetFontAndSize(baseFont, 7);
                cb.SetColorFill(BaseColor.BLACK);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("FROM:"), 10, 275, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("BizLink"), 60, 265, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("NO1.Industrial Zone,FengHuangVillage,"), 12, 255, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("FuYong Town,BaoAn District,ShenZhen"), 12, 245, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("City,GuangDong Province, China"), 12, 235, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("49693A"), 55, 225, 0);

                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("TO:"), 145, 275, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("{0}", shiptos[0]), 150, 265, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("{0}", shiptos[1]), 150, 255, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("{0}", shiptos[2]), 150, 245, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("{0}", shiptos[3]), 150, 235, 0);
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, string.Format("{0}", shiptos[4]), 150, 225, 0);

                cb.EndText();

                //verify position
                //cb.Rectangle(320,135,135,50);
                //cb.SetColorStroke(BaseColor.RED);
                //cb.Stroke();

                ColumnText ct = new ColumnText(cb);
                Phrase p = new Phrase(string.Format("{0}", partDesc), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 12));
                ct.SetSimpleColumn(p, 320, 170, 455, 210, 14, Element.ALIGN_LEFT);
                ct.Go();

                pdf.Close();

                Wait(1);
                axAcroPDF1.setShowToolbar(false);
                if (axAcroPDF1.LoadFile(pdfFile))
                {
                    axAcroPDF1.printAll();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                btnPrint.Enabled = true;
            }
        }
    }
}
