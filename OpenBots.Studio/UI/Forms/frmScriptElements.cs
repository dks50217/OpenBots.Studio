﻿using OpenBots.Core.Script;
using OpenBots.Core.UI.Forms;
using OpenBots.UI.Forms.Supplement_Forms;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace OpenBots.UI.Forms
{
    public partial class frmScriptElements : UIForm
    {
        public ScriptContext ScriptContext { get; set; }
        public string ScriptName { get; set; }
        private TreeNode _userElementParentNode;
        private string _emptyValue = "(no default value)";
        private List<ScriptElement> _elementsCopy;

        #region Initialization and Form Load
        public frmScriptElements()
        {
            InitializeComponent();
        }

        private void frmScriptElements_Load(object sender, EventArgs e)
        {
            //initialize
            _elementsCopy = new List<ScriptElement>(ScriptContext.Elements);
            _elementsCopy = _elementsCopy.OrderBy(x => x.ElementName).ToList();
            _userElementParentNode = InitializeNodes("My Task Elements", _elementsCopy);
            ExpandUserElementNode();
            lblMainLogo.Text = ScriptName + " elements";
        }

        private TreeNode InitializeNodes(string parentName, List<ScriptElement> elements)
        {
            //create a root node (parent)
            TreeNode parentNode = new TreeNode(parentName);

            //add each item to parent
            foreach (var item in elements)
                AddUserElementNode(parentNode, item.ElementName, item.ElementValue);

            //add parent to treeview
            tvScriptElements.Nodes.Add(parentNode);

            //return parent and utilize if needed
            return parentNode;
        }

        #endregion

        #region Add/Cancel Buttons
        private void uiBtnOK_Click(object sender, EventArgs e)
        {
            //return success result
            ScriptContext.Elements = _elementsCopy;
            DialogResult = DialogResult.OK;
        }

        private void uiBtnCancel_Click(object sender, EventArgs e)
        {
            //cancel and close
            DialogResult = DialogResult.Cancel;
        }
        #endregion

        #region Add/Edit Elements
        private void uiBtnNew_Click(object sender, EventArgs e)
        {
            //create element editing form
            frmAddElement addElementForm = new frmAddElement();
            addElementForm.ScriptContext = ScriptContext;
            addElementForm.ElementsCopy = _elementsCopy;

            ExpandUserElementNode();

            //validate if user added element
            if (addElementForm.ShowDialog() == DialogResult.OK)
            {
                //add newly edited node
                AddUserElementNode(_userElementParentNode, addElementForm.txtElementName.Text, addElementForm.ElementValueDT);

                _elementsCopy.Add(new ScriptElement
                {
                    ElementName = addElementForm.txtElementName.Text,
                    ElementValue = addElementForm.ElementValueDT
                });
            }

            addElementForm.Dispose();
        }

        private void tvScriptElements_DoubleClick(object sender, EventArgs e)
        {
            //handle double clicks outside
            if (tvScriptElements.SelectedNode == null)
                return;

            //if parent was selected return
            if (tvScriptElements.SelectedNode.Parent == null)
                return;

            //top node check
            var topNode = GetSelectedTopNode();

            if (topNode.Text != "My Task Elements")
                return;

            ScriptElement element;
            string elementName;
            DataTable elementValue;
            TreeNode parentNode;

            if(tvScriptElements.SelectedNode.Nodes.Count == 0)
            {
                parentNode = tvScriptElements.SelectedNode.Parent;
                elementName = tvScriptElements.SelectedNode.Parent.Text;               
            }
            else
            {
                parentNode = tvScriptElements.SelectedNode;
                elementName = tvScriptElements.SelectedNode.Text;
            }

            element = _elementsCopy.Where(x => x.ElementName == elementName).FirstOrDefault();
            elementValue = element.ElementValue;

            //create element editing form
            frmAddElement addElementForm = new frmAddElement(elementName, elementValue);
            addElementForm.ScriptContext = ScriptContext;
            addElementForm.ElementsCopy = _elementsCopy;

            ExpandUserElementNode();

            //validate if user added element
            if (addElementForm.ShowDialog() == DialogResult.OK)
            {
                //remove parent
                parentNode.Remove();
                AddUserElementNode(_userElementParentNode, addElementForm.txtElementName.Text, addElementForm.ElementValueDT);
            }

            addElementForm.Dispose();
        }

        private void AddUserElementNode(TreeNode parentNode, string elementName, DataTable elementValue)
        {
            //add new node
            var childNode = new TreeNode(elementName);

            for (int i = 0; i < elementValue.Rows.Count; i++)
            {
                if (!string.IsNullOrEmpty(elementValue.Rows[i][2].ToString()))
                {
                    TreeNode elementValueNode = new TreeNode($"ValueNode{i}");
                    string enabled = elementValue.Rows[i][0].ToString();
                    enabled = enabled == "True" ? "Enabled" : "Disabled";
                    elementValueNode.Text = $"{enabled} - {elementValue.Rows[i][1]} - {elementValue.Rows[i][2]}";
                    childNode.Nodes.Add(elementValueNode);
                }               
            }           

            if (childNode.Nodes.Count == 0)
            {
                TreeNode elementValueNode = new TreeNode($"ValueNodeEmpty");
                elementValueNode.Text = _emptyValue;
                childNode.Nodes.Add(elementValueNode);
            }

            parentNode.Nodes.Add(childNode);
            ExpandUserElementNode();
        }

        private void ExpandUserElementNode()
        {
            if (_userElementParentNode != null)
                _userElementParentNode.Expand();
        }

        private void tvScriptElements_KeyDown(object sender, KeyEventArgs e)
        {
            //handling outside
            if (tvScriptElements.SelectedNode == null)
                return;

            //if parent was selected return
            if (tvScriptElements.SelectedNode.Parent == null)
            {
                //user selected top parent
                return;
            }

            //top node check
            var topNode = GetSelectedTopNode();

            if (topNode.Text != "My Task Elements")
                return;

            //if user selected delete
            if (e.KeyCode == Keys.Delete)
            {
                //determine which node is the parent
                TreeNode parentNode;
                if (tvScriptElements.SelectedNode.Nodes.Count == 0)
                    parentNode = tvScriptElements.SelectedNode.Parent;
                else
                    parentNode = tvScriptElements.SelectedNode;

                //remove parent node
                string elementName = parentNode.Text;
                ScriptElement element = _elementsCopy.Where(x => x.ElementName == elementName).FirstOrDefault();
                _elementsCopy.Remove(element);
                parentNode.Remove();
            }
        }

        private TreeNode GetSelectedTopNode()
        {
            TreeNode node = tvScriptElements.SelectedNode;

            while (node.Parent != null)
                node = node.Parent;

            return node;
        }
        #endregion
    }
}