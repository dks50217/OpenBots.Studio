﻿using Microsoft.CodeAnalysis;
using OpenBots.Core.Enums;
using OpenBots.Core.Script;
using OpenBots.Core.UI.Forms;
using OpenBots.Core.Utilities.CommonUtilities;
using OpenBots.Utilities;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace OpenBots.UI.Forms.Supplement_Forms
{
    public partial class frmAddArgument : UIForm
    {
        public ScriptContext ScriptContext { get; set; }
        private bool _isEditMode;
        private string _editingArgumentName;
        private TypeContext _typeContext;
        private Type _preEditType;
        private ToolTip _typeToolTip;
        public List<ScriptArgument> ArgumentsCopy { get; set; }
        private CodeDomProvider _provider;

        public frmAddArgument(TypeContext typeContext)
        {
            InitializeComponent();
            _typeContext = typeContext;
            cbxDefaultDirection.SelectedIndex = 0;

            cbxDefaultType.DataSource = new BindingSource(_typeContext.DefaultTypes, null);
            cbxDefaultType.DisplayMember = "Key";
            cbxDefaultType.ValueMember = "Value";

            cbxDefaultType.SelectedValue = typeof(string);
            cbxDefaultType.Tag = typeof(string);

            _preEditType = typeof(string);
        }

        public frmAddArgument(string argumentName, ScriptArgumentDirection argumentDirection, string argumentValue, Type argumentType,
            TypeContext typeContext)
        {
            InitializeComponent();
            _typeContext = typeContext;
            cbxDefaultType.DataSource = new BindingSource(_typeContext.DefaultTypes, null);
            cbxDefaultType.DisplayMember = "Key";
            cbxDefaultType.ValueMember = "Value";

            Text = "edit argument";
            lblHeader.Text = "edit argument";
            txtArgumentName.Text = argumentName;
            cbxDefaultDirection.Text = argumentDirection.ToString();
            txtDefaultValue.Text = argumentValue;
            cbxDefaultType.SelectedValue = argumentType;
            cbxDefaultType.Tag = argumentType;

            _preEditType = argumentType;

            _isEditMode = true;
            _editingArgumentName = argumentName;
        }

        private void frmAddArgument_Load(object sender, EventArgs e)
        {
            _typeToolTip = AddTypeToolTip();
            _typeToolTip.SetToolTip(cbxDefaultType, _preEditType.GetRealTypeName());
            _provider = CodeDomProvider.CreateProvider("C#");
            ScriptContext.AddIntellisenseControls(Controls);
        }

        private void frmAddArgument_FormClosing(object sender, FormClosingEventArgs e)
        {
            ScriptContext.RemoveIntellisenseControls(Controls);
        }

        private void uiBtnOk_Click(object sender, EventArgs e)
        {
            txtArgumentName.ForeColor = Color.SteelBlue;
            txtDefaultValue.ForeColor = Color.SteelBlue;
            lblArgumentNameError.Text = "";
            lblArgumentValueError.Text = "";

            txtArgumentName.Text = txtArgumentName.Text.Trim();
            if (txtArgumentName.Text == string.Empty)
            {
                lblArgumentNameError.Text = "Argument Name not provided";
                txtArgumentName.ForeColor = Color.Red;
                return;
            }

            if (!_provider.IsValidIdentifier(txtArgumentName.Text))
            {
                lblArgumentNameError.Text = "Argument Name is invalid";
                txtArgumentName.ForeColor = Color.Red;
                return;
            }

            if (!_provider.IsValidIdentifier(txtArgumentName.Text))
            {
                lblArgumentNameError.Text = "Argument Name is invalid";
                return;
            }

            string newArgumentName = txtArgumentName.Text;
            var existingVariable = ScriptContext.Variables.Where(var => var.VariableName == newArgumentName).FirstOrDefault();
            var existingArgument = ArgumentsCopy.Where(var => var.ArgumentName == newArgumentName).FirstOrDefault();
            if (existingArgument != null || existingVariable != null)
            {
                if (!_isEditMode || existingArgument.ArgumentName != _editingArgumentName)
                {
                    lblArgumentNameError.Text = "An Argument or Variable with this name already exists";
                    txtArgumentName.ForeColor = Color.Red;
                    return;
                }
            }

            if (txtArgumentName.Text.StartsWith("{") || txtArgumentName.Text.EndsWith("}"))
            {
                lblArgumentNameError.Text = "Argument markers '{' and '}' should not be included";
                txtArgumentName.ForeColor = Color.Red;
                return;
            }

            var result = ScriptContext.EvaluateVariable(newArgumentName, (Type)cbxDefaultType.Tag, txtDefaultValue.Text);
            if (!result.Success)
            {
                lblArgumentValueError.Text = result.Diagnostics.ToList().Where(x => x.DefaultSeverity == DiagnosticSeverity.Error).FirstOrDefault()?.ToString();
                txtDefaultValue.ForeColor = Color.Red;
                return;
            }

            DialogResult = DialogResult.OK;
        }

        private void uiBtnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void cbxDefaultDirection_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbxDefaultDirection.Text == "Out" || cbxDefaultDirection.Text == "InOut")
            {
                txtDefaultValue.Text = "";
                txtDefaultValue.Enabled = false;
            }
            else
                txtDefaultValue.Enabled = true;               
        }

        private void cbxDefaultDirection_Click(object sender, EventArgs e)
        {
            ComboBox clickedDropdownBox = (ComboBox)sender;
            clickedDropdownBox.DroppedDown = true;
        }

        private void cbxDefaultType_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (((Type)cbxDefaultType.SelectedValue).Name == "MoreOptions")
            {
                frmTypes typeForm = new frmTypes(_typeContext);
                typeForm.ShowDialog();

                if (typeForm.DialogResult == DialogResult.OK)
                {
                    if (!_typeContext.DefaultTypes.ContainsKey(typeForm.SelectedType.GetRealTypeName()))
                    {
                        _typeContext.DefaultTypes.Add(typeForm.SelectedType.GetRealTypeName(), typeForm.SelectedType);
                        cbxDefaultType.DataSource = new BindingSource(_typeContext.DefaultTypes, null);
                    }

                    cbxDefaultType.SelectedValue = typeForm.SelectedType;
                    cbxDefaultType.Tag = typeForm.SelectedType;
                }
                else
                {
                    cbxDefaultType.SelectedValue = _preEditType;
                    cbxDefaultType.Tag = _preEditType;
                }

                typeForm.Dispose();
            }
            else
                cbxDefaultType.Tag = cbxDefaultType.SelectedValue;

            _preEditType = (Type)cbxDefaultType.SelectedValue;
            _typeToolTip.SetToolTip(cbxDefaultType, _preEditType.GetRealTypeName());
        }

        public ToolTip AddTypeToolTip()
        {
            ToolTip typeToolTip = new ToolTip();
            typeToolTip.IsBalloon = false;
            typeToolTip.ShowAlways = true;
            typeToolTip.AutoPopDelay = 5000;
            return typeToolTip;
        }

        private void txtDefaultValue_TextChanged(object sender, EventArgs e)
        {
            ScriptContext.CodeTBXInput_TextChanged(sender, e);
        }

        private void txtDefaultValue_KeyDown(object sender, KeyEventArgs e)
        {
            ScriptContext.CodeInput_KeyDown(sender, e);
        }      
    }
}
