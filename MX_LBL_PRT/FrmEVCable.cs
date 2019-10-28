using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MX_LBL_PRT
{
    public partial class FrmEVCable : Form
    {
        public FrmEVCable()
        {
            InitializeComponent();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            var bmp = ZXingNetBarcodeHelper.GenerateDataMatrix("24-002424-05-01,YYWKSIXXXXXX", 100, 100);
            bmp.Save(@"c:\temp\test.bmp", ImageFormat.Bmp);
        }
    }
}
