using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;
using ZXing;
using ZXing.QrCode;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Microsoft.Reporting.WinForms;




namespace BARC
{
    public partial class Form1 : Form
    {
//        private DataSet1 data = new DataSet1();

        public Form1()
        {
            InitializeComponent();
        }

        private byte[] encoderu(string text, System.Drawing.Imaging.ImageFormat format)
        {
            System.Collections.Generic.Dictionary<EncodeHintType, object> param = new Dictionary<EncodeHintType, object>();
            param.Add(EncodeHintType.CHARACTER_SET, "UTF-8");
            param.Add(EncodeHintType.MARGIN, 0);
            param.Add(EncodeHintType.MIN_SIZE, 9);
            ZXing.QrCode.QRCodeWriter q = new QRCodeWriter();
            return ConvertToBitmapImage(q.encode(text, BarcodeFormat.QR_CODE, 1, 1, param), format);
        }

        private byte[] ConvertToBitmapImage(ZXing.Common.BitMatrix barcode, System.Drawing.Imaging.ImageFormat format)
        {
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            BarcodeWriter bw = new BarcodeWriter();
            Bitmap barcodeBitmap = bw.Write(barcode);
            barcodeBitmap.Save(ms, format);
            ms.Position = 0;
            return ms.ToArray();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Соединяемся с EXCEL файлом и создаем DataSet
            OleDbConnection MyConnection;
            OleDbDataAdapter MyCommand;
            MyConnection = new OleDbConnection(@"provider=Microsoft.Jet.OLEDB.4.0;Data Source='db.xls';Extended Properties=Excel 8.0;");
            MyCommand = new OleDbDataAdapter("select * from [Лист1$]", MyConnection);
            MyCommand.Fill(this.DataSet1.Table1);
            MyConnection.Close();
            string code;
            foreach (BARC.DataSet1.Table1Row row in this.DataSet1.Table1)
            {
                code = row.NAM.Trim();
                try
                {
                    row.BARCODE = encoderu(code, System.Drawing.Imaging.ImageFormat.Bmp);
                }
                catch (Exception err) 
                {
                    if (err == null) { };
                }
            }
            this.reportViewer1.RefreshReport();
        }
    }
}
