using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MX_LBL_PRT
{
    public partial class FrmMain : Form
    {
        public FrmMain()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FrmPolarisLICPlate frmPolarisLICPlate = new FrmPolarisLICPlate();
            frmPolarisLICPlate.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FrmEVCable f = new FrmEVCable();
            f.ShowDialog();
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
            //System.Globalization.Calendar calendar = new System.Globalization.CultureInfo("en-US").Calendar;
            //int weeknum = calendar.GetWeekOfYear(DateTime.Now, System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday);
            ////Dynamic generate button
            //for (int i = 1; i <= 5; i++)
            //{
            //    Button button = new Button();
            //    button.Size = new Size(120, 60);
            //    button.Location = new Point(10, (i - 1) * 60 + 10 * i);
            //    button.Text = "Test " + i.ToString();
            //    button.Tag = weeknum;
            //    button.Click += new System.EventHandler(this.buttons_Click);
            //    panel1.Controls.Add(button);
            //}
            //
        }

        private void buttons_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            MessageBox.Show(button.Text);
            MessageBox.Show(button.Tag.ToString());
        }
    }
}
