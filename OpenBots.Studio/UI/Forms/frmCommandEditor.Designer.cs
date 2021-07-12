﻿namespace OpenBots.UI.Forms
{
    partial class frmCommandEditor
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmCommandEditor));
            this.cboSelectedCommand = new System.Windows.Forms.ComboBox();
            this.flw_InputVariables = new System.Windows.Forms.FlowLayoutPanel();
            this.tlpCommandControls = new System.Windows.Forms.TableLayoutPanel();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.uiBtnAdd = new OpenBots.Core.UI.Controls.UIPictureButton();
            this.uiBtnCancel = new OpenBots.Core.UI.Controls.UIPictureButton();
            this.tlpCommandControls.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.uiBtnAdd)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.uiBtnCancel)).BeginInit();
            this.SuspendLayout();
            // 
            // cboSelectedCommand
            // 
            this.cboSelectedCommand.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboSelectedCommand.BackColor = System.Drawing.Color.WhiteSmoke;
            this.cboSelectedCommand.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboSelectedCommand.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.cboSelectedCommand.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cboSelectedCommand.FormattingEnabled = true;
            this.cboSelectedCommand.Location = new System.Drawing.Point(6, 5);
            this.cboSelectedCommand.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.cboSelectedCommand.Name = "cboSelectedCommand";
            this.cboSelectedCommand.Size = new System.Drawing.Size(619, 33);
            this.cboSelectedCommand.TabIndex = 2;
            this.cboSelectedCommand.SelectionChangeCommitted += new System.EventHandler(this.cboSelectedCommand_SelectionChangeCommitted);
            this.cboSelectedCommand.MouseWheel += new System.Windows.Forms.MouseEventHandler(this.cboSelectedCommand_MouseWheel);
            // 
            // flw_InputVariables
            // 
            this.flw_InputVariables.AutoScroll = true;
            this.flw_InputVariables.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.flw_InputVariables.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(59)))), ((int)(((byte)(59)))), ((int)(((byte)(59)))));
            this.flw_InputVariables.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flw_InputVariables.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.flw_InputVariables.Font = new System.Drawing.Font("Segoe UI Semibold", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.flw_InputVariables.Location = new System.Drawing.Point(5, 38);
            this.flw_InputVariables.Margin = new System.Windows.Forms.Padding(5);
            this.flw_InputVariables.Name = "flw_InputVariables";
            this.flw_InputVariables.Padding = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.flw_InputVariables.Size = new System.Drawing.Size(621, 645);
            this.flw_InputVariables.TabIndex = 3;
            this.flw_InputVariables.WrapContents = false;
            // 
            // tlpCommandControls
            // 
            this.tlpCommandControls.BackColor = System.Drawing.Color.Transparent;
            this.tlpCommandControls.ColumnCount = 1;
            this.tlpCommandControls.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tlpCommandControls.Controls.Add(this.cboSelectedCommand, 0, 0);
            this.tlpCommandControls.Controls.Add(this.flw_InputVariables, 0, 1);
            this.tlpCommandControls.Controls.Add(this.flowLayoutPanel1, 0, 2);
            this.tlpCommandControls.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tlpCommandControls.Location = new System.Drawing.Point(0, 0);
            this.tlpCommandControls.Name = "tlpCommandControls";
            this.tlpCommandControls.RowCount = 3;
            this.tlpCommandControls.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 33F));
            this.tlpCommandControls.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tlpCommandControls.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 73F));
            this.tlpCommandControls.Size = new System.Drawing.Size(631, 761);
            this.tlpCommandControls.TabIndex = 17;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.BackColor = System.Drawing.Color.Transparent;
            this.flowLayoutPanel1.Controls.Add(this.uiBtnAdd);
            this.flowLayoutPanel1.Controls.Add(this.uiBtnCancel);
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(0, 688);
            this.flowLayoutPanel1.Margin = new System.Windows.Forms.Padding(0);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(631, 73);
            this.flowLayoutPanel1.TabIndex = 4;
            // 
            // uiBtnAdd
            // 
            this.uiBtnAdd.BackColor = System.Drawing.Color.Transparent;
            this.uiBtnAdd.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.uiBtnAdd.DisplayText = "Ok";
            this.uiBtnAdd.DisplayTextBrush = System.Drawing.Color.White;
            this.uiBtnAdd.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.uiBtnAdd.Image = ((System.Drawing.Image)(resources.GetObject("uiBtnAdd.Image")));
            this.uiBtnAdd.IsMouseOver = false;
            this.uiBtnAdd.Location = new System.Drawing.Point(6, 5);
            this.uiBtnAdd.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.uiBtnAdd.Name = "uiBtnAdd";
            this.uiBtnAdd.Size = new System.Drawing.Size(60, 61);
            this.uiBtnAdd.TabIndex = 14;
            this.uiBtnAdd.TabStop = false;
            this.uiBtnAdd.Text = "Ok";
            this.uiBtnAdd.Click += new System.EventHandler(this.uiBtnAdd_Click);
            // 
            // uiBtnCancel
            // 
            this.uiBtnCancel.BackColor = System.Drawing.Color.Transparent;
            this.uiBtnCancel.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.uiBtnCancel.DisplayText = "Cancel";
            this.uiBtnCancel.DisplayTextBrush = System.Drawing.Color.White;
            this.uiBtnCancel.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.uiBtnCancel.Image = ((System.Drawing.Image)(resources.GetObject("uiBtnCancel.Image")));
            this.uiBtnCancel.IsMouseOver = false;
            this.uiBtnCancel.Location = new System.Drawing.Point(78, 5);
            this.uiBtnCancel.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.uiBtnCancel.Name = "uiBtnCancel";
            this.uiBtnCancel.Size = new System.Drawing.Size(60, 61);
            this.uiBtnCancel.TabIndex = 15;
            this.uiBtnCancel.TabStop = false;
            this.uiBtnCancel.Text = "Cancel";
            this.uiBtnCancel.Click += new System.EventHandler(this.uiBtnCancel_Click);
            // 
            // frmCommandEditor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 29F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SteelBlue;
            this.ClientSize = new System.Drawing.Size(631, 761);
            this.Controls.Add(this.tlpCommandControls);
            this.DoubleBuffered = true;
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(6, 5, 6, 5);
            this.Name = "frmCommandEditor";
            this.Text = "Add New Command";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmCommandEditor_FormClosing);
            this.Load += new System.EventHandler(this.frmNewCommand_Load);
            this.Shown += new System.EventHandler(this.frmCommandEditor_Shown);
            this.Resize += new System.EventHandler(this.frmCommandEditor_Resize);
            this.tlpCommandControls.ResumeLayout(false);
            this.flowLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.uiBtnAdd)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.uiBtnCancel)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.ComboBox cboSelectedCommand;
        private OpenBots.Core.UI.Controls.UIPictureButton uiBtnCancel;
        private OpenBots.Core.UI.Controls.UIPictureButton uiBtnAdd;
        private System.Windows.Forms.TableLayoutPanel tlpCommandControls;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;        
    }
}