using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace Chudisoft.InputControls
{
    public partial class DateControlMasked : UserControl
    {
        public DateTimePickerBorderless dtp = new DateTimePickerBorderless();

        public DateControlMasked()
        {
            InitializeComponent();
            panel1.Controls.Add(dtp);
            //this.Size = new Size(250, 24);
            //dtp.Width = this.Width - 2;
            dtp.Location = new Point(dtp.Location.X + 2, dtp.Location.Y + 2);
            dtp.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            dtp.Format = DateTimePickerFormat.Custom;
            dtp.CustomFormat = "dd-MM-yyyy";
            dtp.Visible = false;

            toolTip1.SetToolTip(dtp, toolTip1.GetToolTip(this));
        }

        private void mtxt_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(mtxt.Text))
                {
                    dtp.Value = DateTime.ParseExact(mtxt.Text,
                    "dd-MM-yyyy", System.Globalization.CultureInfo.InvariantCulture);
                }
            }
            catch (Exception ex)
            {
            }
        }

        internal void SetText(string DateString)
        {
            try
            {
                if (!string.IsNullOrEmpty(DateString))
                {
                    mtxt.Text = DateString;
                }
            }
            catch (Exception ex)
            {
            }
        }
    }
}
