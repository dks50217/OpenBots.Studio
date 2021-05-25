﻿namespace OpenBots.UI.Forms.Supplement_Forms
{
    partial class frmAddArgument
    {
        /// <summary>
        /// Required designer argument.
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmAddArgument));
            this.lblDefineName = new System.Windows.Forms.Label();
            this.lblHeader = new System.Windows.Forms.Label();
            this.txtArgumentName = new System.Windows.Forms.TextBox();
            this.lblDefineNameDescription = new System.Windows.Forms.Label();
            this.lblDefineDefaultValueDescriptor = new System.Windows.Forms.Label();
            this.txtDefaultValue = new System.Windows.Forms.TextBox();
            this.lblDefineDefaultValue = new System.Windows.Forms.Label();
            this.uiBtnOk = new OpenBots.Core.UI.Controls.UIPictureButton();
            this.uiBtnCancel = new OpenBots.Core.UI.Controls.UIPictureButton();
            this.lblArgumentNameError = new System.Windows.Forms.Label();
            this.lblDefineDefaultDirectionDescriptor = new System.Windows.Forms.Label();
            this.cbxDefaultDirection = new System.Windows.Forms.ComboBox();
            this.lblDefineDefaultDirection = new System.Windows.Forms.Label();
            this.lblDefineDefaultTypeDescriptor = new System.Windows.Forms.Label();
            this.cbxDefaultType = new System.Windows.Forms.ComboBox();
            this.lblDefineDefaultType = new System.Windows.Forms.Label();
            this.lblArgumentValueError = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.uiBtnOk)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.uiBtnCancel)).BeginInit();
            this.SuspendLayout();
            // 
            // lblDefineName
            // 
            this.lblDefineName.AutoSize = true;
            this.lblDefineName.BackColor = System.Drawing.Color.Transparent;
            this.lblDefineName.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDefineName.ForeColor = System.Drawing.Color.White;
            this.lblDefineName.Location = new System.Drawing.Point(16, 60);
            this.lblDefineName.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblDefineName.Name = "lblDefineName";
            this.lblDefineName.Size = new System.Drawing.Size(230, 28);
            this.lblDefineName.TabIndex = 15;
            this.lblDefineName.Text = "Define Argument Name";
            // 
            // lblHeader
            // 
            this.lblHeader.AutoSize = true;
            this.lblHeader.BackColor = System.Drawing.Color.Transparent;
            this.lblHeader.Font = new System.Drawing.Font("Segoe UI Semilight", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHeader.ForeColor = System.Drawing.Color.WhiteSmoke;
            this.lblHeader.Location = new System.Drawing.Point(8, 4);
            this.lblHeader.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblHeader.Name = "lblHeader";
            this.lblHeader.Size = new System.Drawing.Size(268, 54);
            this.lblHeader.TabIndex = 14;
            this.lblHeader.Text = "add argument";
            // 
            // txtArgumentName
            // 
            this.txtArgumentName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtArgumentName.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtArgumentName.ForeColor = System.Drawing.Color.SteelBlue;
            this.txtArgumentName.Location = new System.Drawing.Point(21, 168);
            this.txtArgumentName.Margin = new System.Windows.Forms.Padding(4);
            this.txtArgumentName.Name = "txtArgumentName";
            this.txtArgumentName.Size = new System.Drawing.Size(566, 32);
            this.txtArgumentName.TabIndex = 16;
            // 
            // lblDefineNameDescription
            // 
            this.lblDefineNameDescription.BackColor = System.Drawing.Color.Transparent;
            this.lblDefineNameDescription.Font = new System.Drawing.Font("Segoe UI Light", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDefineNameDescription.ForeColor = System.Drawing.Color.WhiteSmoke;
            this.lblDefineNameDescription.Location = new System.Drawing.Point(16, 86);
            this.lblDefineNameDescription.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblDefineNameDescription.Name = "lblDefineNameDescription";
            this.lblDefineNameDescription.Size = new System.Drawing.Size(571, 79);
            this.lblDefineNameDescription.TabIndex = 17;
            this.lblDefineNameDescription.Text = "Define a name for your argument, such as \'in_Number\'. Remember to enclose the nam" +
    "e within brackets in order to use the argument in commands.";
            // 
            // lblDefineDefaultValueDescriptor
            // 
            this.lblDefineDefaultValueDescriptor.BackColor = System.Drawing.Color.Transparent;
            this.lblDefineDefaultValueDescriptor.Font = new System.Drawing.Font("Segoe UI Light", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDefineDefaultValueDescriptor.ForeColor = System.Drawing.Color.WhiteSmoke;
            this.lblDefineDefaultValueDescriptor.Location = new System.Drawing.Point(16, 556);
            this.lblDefineDefaultValueDescriptor.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblDefineDefaultValueDescriptor.Name = "lblDefineDefaultValueDescriptor";
            this.lblDefineDefaultValueDescriptor.Size = new System.Drawing.Size(571, 53);
            this.lblDefineDefaultValueDescriptor.TabIndex = 20;
            this.lblDefineDefaultValueDescriptor.Text = "Optionally, define a default value for the argument. The argument will represent " +
    "this value until changed during the task by a task command.";
            // 
            // txtDefaultValue
            // 
            this.txtDefaultValue.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDefaultValue.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDefaultValue.ForeColor = System.Drawing.Color.SteelBlue;
            this.txtDefaultValue.Location = new System.Drawing.Point(21, 613);
            this.txtDefaultValue.Margin = new System.Windows.Forms.Padding(4);
            this.txtDefaultValue.Name = "txtDefaultValue";
            this.txtDefaultValue.Size = new System.Drawing.Size(566, 32);
            this.txtDefaultValue.TabIndex = 19;
            // 
            // lblDefineDefaultValue
            // 
            this.lblDefineDefaultValue.AutoSize = true;
            this.lblDefineDefaultValue.BackColor = System.Drawing.Color.Transparent;
            this.lblDefineDefaultValue.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDefineDefaultValue.ForeColor = System.Drawing.Color.White;
            this.lblDefineDefaultValue.Location = new System.Drawing.Point(16, 530);
            this.lblDefineDefaultValue.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblDefineDefaultValue.Name = "lblDefineDefaultValue";
            this.lblDefineDefaultValue.Size = new System.Drawing.Size(298, 28);
            this.lblDefineDefaultValue.TabIndex = 18;
            this.lblDefineDefaultValue.Text = "Define Argument Default Value";
            // 
            // uiBtnOk
            // 
            this.uiBtnOk.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.uiBtnOk.BackColor = System.Drawing.Color.Transparent;
            this.uiBtnOk.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.uiBtnOk.DisplayText = "Ok";
            this.uiBtnOk.DisplayTextBrush = System.Drawing.Color.White;
            this.uiBtnOk.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.uiBtnOk.Image = ((System.Drawing.Image)(resources.GetObject("uiBtnOk.Image")));
            this.uiBtnOk.IsMouseOver = false;
            this.uiBtnOk.Location = new System.Drawing.Point(10, 695);
            this.uiBtnOk.Margin = new System.Windows.Forms.Padding(8, 6, 8, 6);
            this.uiBtnOk.Name = "uiBtnOk";
            this.uiBtnOk.Size = new System.Drawing.Size(60, 60);
            this.uiBtnOk.TabIndex = 21;
            this.uiBtnOk.TabStop = false;
            this.uiBtnOk.Text = "Ok";
            this.uiBtnOk.Click += new System.EventHandler(this.uiBtnOk_Click);
            // 
            // uiBtnCancel
            // 
            this.uiBtnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.uiBtnCancel.BackColor = System.Drawing.Color.Transparent;
            this.uiBtnCancel.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.uiBtnCancel.DisplayText = "Cancel";
            this.uiBtnCancel.DisplayTextBrush = System.Drawing.Color.White;
            this.uiBtnCancel.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.uiBtnCancel.Image = ((System.Drawing.Image)(resources.GetObject("uiBtnCancel.Image")));
            this.uiBtnCancel.IsMouseOver = false;
            this.uiBtnCancel.Location = new System.Drawing.Point(70, 695);
            this.uiBtnCancel.Margin = new System.Windows.Forms.Padding(8, 6, 8, 6);
            this.uiBtnCancel.Name = "uiBtnCancel";
            this.uiBtnCancel.Size = new System.Drawing.Size(60, 60);
            this.uiBtnCancel.TabIndex = 22;
            this.uiBtnCancel.TabStop = false;
            this.uiBtnCancel.Text = "Cancel";
            this.uiBtnCancel.Click += new System.EventHandler(this.uiBtnCancel_Click);
            // 
            // lblArgumentNameError
            // 
            this.lblArgumentNameError.BackColor = System.Drawing.Color.Transparent;
            this.lblArgumentNameError.Font = new System.Drawing.Font("Segoe UI Light", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblArgumentNameError.ForeColor = System.Drawing.Color.Red;
            this.lblArgumentNameError.Location = new System.Drawing.Point(16, 200);
            this.lblArgumentNameError.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblArgumentNameError.Name = "lblArgumentNameError";
            this.lblArgumentNameError.Size = new System.Drawing.Size(571, 36);
            this.lblArgumentNameError.TabIndex = 23;
            // 
            // lblDefineDefaultDirectionDescriptor
            // 
            this.lblDefineDefaultDirectionDescriptor.BackColor = System.Drawing.Color.Transparent;
            this.lblDefineDefaultDirectionDescriptor.Font = new System.Drawing.Font("Segoe UI Light", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDefineDefaultDirectionDescriptor.ForeColor = System.Drawing.Color.WhiteSmoke;
            this.lblDefineDefaultDirectionDescriptor.Location = new System.Drawing.Point(16, 406);
            this.lblDefineDefaultDirectionDescriptor.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblDefineDefaultDirectionDescriptor.Name = "lblDefineDefaultDirectionDescriptor";
            this.lblDefineDefaultDirectionDescriptor.Size = new System.Drawing.Size(571, 53);
            this.lblDefineDefaultDirectionDescriptor.TabIndex = 26;
            this.lblDefineDefaultDirectionDescriptor.Text = "Define a default direction for the argument. Arguments with Out Directions do not" +
    " support the assignment of default values.";
            // 
            // cbxDefaultDirection
            // 
            this.cbxDefaultDirection.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cbxDefaultDirection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxDefaultDirection.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbxDefaultDirection.ForeColor = System.Drawing.Color.SteelBlue;
            this.cbxDefaultDirection.Items.AddRange(new object[] {
            "In",
            "Out",
            "InOut"});
            this.cbxDefaultDirection.Location = new System.Drawing.Point(21, 463);
            this.cbxDefaultDirection.Margin = new System.Windows.Forms.Padding(4);
            this.cbxDefaultDirection.Name = "cbxDefaultDirection";
            this.cbxDefaultDirection.Size = new System.Drawing.Size(566, 33);
            this.cbxDefaultDirection.TabIndex = 25;
            this.cbxDefaultDirection.SelectedIndexChanged += new System.EventHandler(this.cbxDefaultDirection_SelectedIndexChanged);
            this.cbxDefaultDirection.Click += new System.EventHandler(this.cbxDefaultDirection_Click);
            // 
            // lblDefineDefaultDirection
            // 
            this.lblDefineDefaultDirection.AutoSize = true;
            this.lblDefineDefaultDirection.BackColor = System.Drawing.Color.Transparent;
            this.lblDefineDefaultDirection.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDefineDefaultDirection.ForeColor = System.Drawing.Color.White;
            this.lblDefineDefaultDirection.Location = new System.Drawing.Point(16, 380);
            this.lblDefineDefaultDirection.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblDefineDefaultDirection.Name = "lblDefineDefaultDirection";
            this.lblDefineDefaultDirection.Size = new System.Drawing.Size(330, 28);
            this.lblDefineDefaultDirection.TabIndex = 24;
            this.lblDefineDefaultDirection.Text = "Define Argument Default Direction";
            // 
            // lblDefineDefaultTypeDescriptor
            // 
            this.lblDefineDefaultTypeDescriptor.BackColor = System.Drawing.Color.Transparent;
            this.lblDefineDefaultTypeDescriptor.Font = new System.Drawing.Font("Segoe UI Light", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDefineDefaultTypeDescriptor.ForeColor = System.Drawing.Color.WhiteSmoke;
            this.lblDefineDefaultTypeDescriptor.Location = new System.Drawing.Point(16, 266);
            this.lblDefineDefaultTypeDescriptor.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblDefineDefaultTypeDescriptor.Name = "lblDefineDefaultTypeDescriptor";
            this.lblDefineDefaultTypeDescriptor.Size = new System.Drawing.Size(571, 53);
            this.lblDefineDefaultTypeDescriptor.TabIndex = 29;
            this.lblDefineDefaultTypeDescriptor.Text = "Define a default type for the argument. The type of the argument cannot be change" +
    "d during the execution of a Script.";
            // 
            // cbxDefaultType
            // 
            this.cbxDefaultType.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cbxDefaultType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxDefaultType.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbxDefaultType.ForeColor = System.Drawing.Color.SteelBlue;
            this.cbxDefaultType.Location = new System.Drawing.Point(21, 323);
            this.cbxDefaultType.Margin = new System.Windows.Forms.Padding(4);
            this.cbxDefaultType.Name = "cbxDefaultType";
            this.cbxDefaultType.Size = new System.Drawing.Size(566, 33);
            this.cbxDefaultType.TabIndex = 28;
            this.cbxDefaultType.Tag = typeof(string);
            this.cbxDefaultType.SelectionChangeCommitted += new System.EventHandler(this.cbxDefaultType_SelectionChangeCommitted);
            // 
            // lblDefineDefaultType
            // 
            this.lblDefineDefaultType.AutoSize = true;
            this.lblDefineDefaultType.BackColor = System.Drawing.Color.Transparent;
            this.lblDefineDefaultType.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDefineDefaultType.ForeColor = System.Drawing.Color.White;
            this.lblDefineDefaultType.Location = new System.Drawing.Point(16, 240);
            this.lblDefineDefaultType.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblDefineDefaultType.Name = "lblDefineDefaultType";
            this.lblDefineDefaultType.Size = new System.Drawing.Size(291, 28);
            this.lblDefineDefaultType.TabIndex = 27;
            this.lblDefineDefaultType.Text = "Define Argument Default Type";
            // 
            // lblArgumentValueError
            // 
            this.lblArgumentValueError.BackColor = System.Drawing.Color.Transparent;
            this.lblArgumentValueError.Font = new System.Drawing.Font("Segoe UI Light", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblArgumentValueError.ForeColor = System.Drawing.Color.Red;
            this.lblArgumentValueError.Location = new System.Drawing.Point(16, 645);
            this.lblArgumentValueError.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblArgumentValueError.Name = "lblArgumentValueError";
            this.lblArgumentValueError.Size = new System.Drawing.Size(571, 36);
            this.lblArgumentValueError.TabIndex = 34;
            // 
            // frmAddArgument
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(609, 769);
            this.Controls.Add(this.lblArgumentValueError);
            this.Controls.Add(this.lblDefineDefaultTypeDescriptor);
            this.Controls.Add(this.cbxDefaultType);
            this.Controls.Add(this.lblDefineDefaultType);
            this.Controls.Add(this.lblDefineDefaultDirectionDescriptor);
            this.Controls.Add(this.cbxDefaultDirection);
            this.Controls.Add(this.lblDefineDefaultDirection);
            this.Controls.Add(this.lblArgumentNameError);
            this.Controls.Add(this.uiBtnOk);
            this.Controls.Add(this.uiBtnCancel);
            this.Controls.Add(this.lblDefineDefaultValueDescriptor);
            this.Controls.Add(this.txtDefaultValue);
            this.Controls.Add(this.lblDefineDefaultValue);
            this.Controls.Add(this.lblDefineNameDescription);
            this.Controls.Add(this.txtArgumentName);
            this.Controls.Add(this.lblDefineName);
            this.Controls.Add(this.lblHeader);
            this.Icon = global::OpenBots.Properties.Resources.OpenBots_ico;
            this.Margin = new System.Windows.Forms.Padding(5);
            this.MinimumSize = new System.Drawing.Size(627, 491);
            this.Name = "frmAddArgument";
            this.Text = "Add Argument";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.frmAddArgument_Load);
            ((System.ComponentModel.ISupportInitialize)(this.uiBtnOk)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.uiBtnCancel)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblDefineName;
        private System.Windows.Forms.Label lblHeader;
        private System.Windows.Forms.Label lblDefineNameDescription;
        private System.Windows.Forms.Label lblDefineDefaultValueDescriptor;
        private System.Windows.Forms.Label lblDefineDefaultValue;
        private OpenBots.Core.UI.Controls.UIPictureButton uiBtnOk;
        private OpenBots.Core.UI.Controls.UIPictureButton uiBtnCancel;
        public System.Windows.Forms.TextBox txtArgumentName;
        public System.Windows.Forms.TextBox txtDefaultValue;
        private System.Windows.Forms.Label lblArgumentNameError;
        private System.Windows.Forms.Label lblDefineDefaultDirectionDescriptor;
        public System.Windows.Forms.ComboBox cbxDefaultDirection;
        private System.Windows.Forms.Label lblDefineDefaultDirection;
        private System.Windows.Forms.Label lblDefineDefaultTypeDescriptor;
        public System.Windows.Forms.ComboBox cbxDefaultType;
        private System.Windows.Forms.Label lblDefineDefaultType;
        private System.Windows.Forms.Label lblArgumentValueError;
    }
}