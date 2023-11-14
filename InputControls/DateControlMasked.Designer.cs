using System.Windows.Forms;

namespace Chudisoft.InputControls
{
    partial class DateControlMasked
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.panel1 = new System.Windows.Forms.Panel();
            this.mtxt = new System.Windows.Forms.MaskedTextBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.mtxt);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(82, 16);
            this.panel1.TabIndex = 0;
            // 
            // mtxt
            // 
            this.mtxt.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.mtxt.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mtxt.Enabled = false;
            this.mtxt.HidePromptOnLeave = true;
            this.mtxt.Location = new System.Drawing.Point(0, 0);
            this.mtxt.Mask = "00-00-0000";
            this.mtxt.Name = "mtxt";
            this.mtxt.ReadOnly = true;
            this.mtxt.Size = new System.Drawing.Size(82, 16);
            this.mtxt.TabIndex = 1;
            this.mtxt.TabStop = false;
            this.mtxt.ValidatingType = typeof(System.DateTime);
            this.mtxt.TextChanged += new System.EventHandler(this.mtxt_TextChanged);
            // 
            // toolTip1
            // 
            this.toolTip1.IsBalloon = true;
            this.toolTip1.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            // 
            // DateControlMasked
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.Transparent;
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.Margin = new System.Windows.Forms.Padding(1);
            this.MaximumSize = new System.Drawing.Size(255, 30);
            this.Name = "DateControlMasked";
            this.Padding = new System.Windows.Forms.Padding(3);
            this.Size = new System.Drawing.Size(88, 22);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private Panel panel1;
        private MaskedTextBox mtxt;
        private ToolTip toolTip1;
    }
}
