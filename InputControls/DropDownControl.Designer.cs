using System;
using System.Drawing;
using System.Windows.Forms;

namespace Chudisoft.InputControls
{
    partial class DropDownControl
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
            this.cmbList = new System.Windows.Forms.ComboBox();
            this.lblText = new System.Windows.Forms.Label();
            this.btnIcon = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // cmbList
            // 
            this.cmbList.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(230)))), ((int)(((byte)(228)))), ((int)(((byte)(245)))));
            this.cmbList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cmbList.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.cmbList.ForeColor = System.Drawing.Color.DimGray;
            this.cmbList.Location = new System.Drawing.Point(0, 0);
            this.cmbList.Name = "cmbList";
            this.cmbList.Size = new System.Drawing.Size(198, 24);
            this.cmbList.TabIndex = 0;
            this.cmbList.SelectedIndexChanged += new System.EventHandler(this.ComboBox_SelectedIndexChanged);
            this.cmbList.TextChanged += new System.EventHandler(this.ComboBox_TextChanged);
            this.cmbList.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbList_KeyPress);
            // 
            // lblText
            // 
            this.lblText.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblText.BackColor = System.Drawing.Color.WhiteSmoke;
            this.lblText.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.lblText.Location = new System.Drawing.Point(1, 2);
            this.lblText.Name = "lblText";
            this.lblText.Padding = new System.Windows.Forms.Padding(8, 0, 0, 0);
            this.lblText.Size = new System.Drawing.Size(168, 20);
            this.lblText.TabIndex = 0;
            this.lblText.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblText.Click += new System.EventHandler(this.Surface_Click);
            this.lblText.MouseEnter += new System.EventHandler(this.Surface_MouseEnter);
            this.lblText.MouseLeave += new System.EventHandler(this.Surface_MouseLeave);
            // 
            // btnIcon
            // 
            this.btnIcon.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnIcon.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btnIcon.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnIcon.FlatAppearance.BorderSize = 0;
            this.btnIcon.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnIcon.Location = new System.Drawing.Point(169, 2);
            this.btnIcon.Name = "btnIcon";
            this.btnIcon.Size = new System.Drawing.Size(30, 20);
            this.btnIcon.TabIndex = 1;
            this.btnIcon.TabStop = false;
            this.btnIcon.UseVisualStyleBackColor = false;
            this.btnIcon.Click += new System.EventHandler(this.Icon_Click);
            this.btnIcon.Paint += new System.Windows.Forms.PaintEventHandler(this.Icon_Paint);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.cmbList);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(1, 1);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(198, 23);
            this.panel1.TabIndex = 0;
            // 
            // toolTip1
            // 
            this.toolTip1.IsBalloon = true;
            this.toolTip1.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            // 
            // DropDownControl
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.Controls.Add(this.lblText);
            this.Controls.Add(this.btnIcon);
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.ForeColor = System.Drawing.Color.DimGray;
            this.Margin = new System.Windows.Forms.Padding(1);
            this.MinimumSize = new System.Drawing.Size(200, 20);
            this.Name = "DropDownControl";
            this.Padding = new System.Windows.Forms.Padding(1);
            this.Size = new System.Drawing.Size(200, 25);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        //Items
        public ComboBox cmbList;
        private Label lblText;
        private Button btnIcon;
        private Panel panel1;
        private ToolTip toolTip1;
    }
}
