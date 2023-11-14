using System;
using System.Drawing;
using System.Windows.Forms;

namespace Chudisoft.InputControls
{
    partial class DropDownControlText
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
            this.txt = new System.Windows.Forms.TextBox();
            this.btnIcon = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.SuspendLayout();
            // 
            // cmbList
            // 
            this.cmbList.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbList.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(230)))), ((int)(((byte)(228)))), ((int)(((byte)(245)))));
            this.cmbList.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cmbList.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.5F);
            this.cmbList.ForeColor = System.Drawing.Color.DimGray;
            this.cmbList.Location = new System.Drawing.Point(0, 0);
            this.cmbList.Name = "cmbList";
            this.cmbList.Size = new System.Drawing.Size(170, 24);
            this.cmbList.TabIndex = 2;
            this.cmbList.SelectedIndexChanged += new System.EventHandler(this.ComboBox_SelectedIndexChanged);
            this.cmbList.TextChanged += new System.EventHandler(this.ComboBox_TextChanged);
            this.cmbList.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbList_KeyPress);
            // 
            // txt
            // 
            this.txt.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txt.BackColor = System.Drawing.Color.WhiteSmoke;
            this.txt.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txt.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.5F);
            this.txt.Location = new System.Drawing.Point(1, 4);
            this.txt.Name = "txt";
            this.txt.Size = new System.Drawing.Size(175, 15);
            this.txt.TabIndex = 0;
            this.txt.Click += new System.EventHandler(this.Surface_Click);
            this.txt.MouseEnter += new System.EventHandler(this.Surface_MouseEnter);
            this.txt.MouseLeave += new System.EventHandler(this.Surface_MouseLeave);
            // 
            // btnIcon
            // 
            this.btnIcon.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnIcon.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btnIcon.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnIcon.FlatAppearance.BorderSize = 0;
            this.btnIcon.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnIcon.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.5F);
            this.btnIcon.Location = new System.Drawing.Point(176, 4);
            this.btnIcon.Name = "btnIcon";
            this.btnIcon.Size = new System.Drawing.Size(30, 23);
            this.btnIcon.TabIndex = 1;
            this.btnIcon.TabStop = false;
            this.btnIcon.UseVisualStyleBackColor = false;
            this.btnIcon.Click += new System.EventHandler(this.Icon_Click);
            this.btnIcon.Paint += new System.Windows.Forms.PaintEventHandler(this.Icon_Paint);
            // 
            // toolTip1
            // 
            this.toolTip1.IsBalloon = true;
            this.toolTip1.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            // 
            // DropDownControlText
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.White;
            this.Controls.Add(this.txt);
            this.Controls.Add(this.btnIcon);
            this.Controls.Add(this.cmbList);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.5F);
            this.ForeColor = System.Drawing.Color.DimGray;
            this.Margin = new System.Windows.Forms.Padding(1);
            this.MinimumSize = new System.Drawing.Size(200, 20);
            this.Name = "DropDownControlText";
            this.Padding = new System.Windows.Forms.Padding(1);
            this.Size = new System.Drawing.Size(207, 28);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        //Items
        public ComboBox cmbList;
        private TextBox txt;
        private Button btnIcon;
        private ToolTip toolTip1;
    }
}
