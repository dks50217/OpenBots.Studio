﻿using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Interfaces;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace OpenBots.Core.UI.Controls
{
    public class AutomationCommand
    {
        public Type CommandClass { get; set; }
        public string FullName { get; set; }
        public string ShortName { get; set; }
        public string Description { get; set; }
        public string DisplayGroup { get; set; }
        public ScriptCommand Command { get; set; }
        public List<Control> UIControls { get; set; }

        private void RenderUIComponents(IfrmCommandEditor editorForm, ICommandControls commandControls)
        {
            if (Command == null)
            {
                throw new InvalidOperationException("Command cannot be null!");
            }

            UIControls = new List<Control>();

            var renderedControls = Command.Render(editorForm, commandControls);

            foreach (var ctrl in renderedControls)
            {
                UIControls.Add(ctrl);
            }

            //generate Private Checkbox (Control) if user did not add it
            var privateCheckBoxControlExists = renderedControls.Any(f => f.Name == "v_IsPrivate");
            var continueComboBoxControlExists = renderedControls.Any(f => f.Name == "v_ErrorHandling");

            if (!privateCheckBoxControlExists)
            {
                //TODO: when using a layoutpanel, checkbox is resetting when form closes
                //FlowLayoutPanel flpCheckBox = new FlowLayoutPanel();
                //flpCheckBox.Height = 30;
                //flpCheckBox.FlowDirection = FlowDirection.LeftToRight;
                UIControls.Add(commandControls.CreateDefaultLabelFor("v_IsPrivate", Command));
                UIControls.Add(commandControls.CreateCheckBoxFor("v_IsPrivate", Command));
                //UIControls.Add(flpCheckBox);
            }

            if (!continueComboBoxControlExists)
            {
                UIControls.Add(commandControls.CreateDefaultLabelFor("v_ErrorHandling", Command));
                ComboBox continueBox = commandControls.CreateDropdownFor("v_ErrorHandling", Command);
                continueBox.Items.Add("None");
                continueBox.Items.Add("Report Error");
                continueBox.Items.Add("Ignore Error");
                UIControls.Add(continueBox);
            }

            //generate comment command if user did not generate it
            var commentControlExists = renderedControls.Any(f => f.Name == "v_Comment");

            if (!commentControlExists)
            {
                UIControls.Add(commandControls.CreateDefaultLabelFor("v_Comment", Command));
                UIControls.Add(commandControls.CreateDefaultNoValidationInputFor("v_Comment", Command, 100, 300));
            }
        }

        public void Bind(IfrmCommandEditor editor, ICommandControls commandControls)
        {
            RenderUIComponents(editor, commandControls);

            foreach (var ctrl in UIControls)
            {
                if (ctrl.DataBindings.Count > 0)
                {
                    var newBindingList = new List<Binding>();
                    foreach (Binding binding in ctrl.DataBindings)
                    {
                        newBindingList.Add(
                            new Binding(
                                binding.PropertyName,
                                Command,
                                binding.BindingMemberInfo.BindingField,
                                false,
                                DataSourceUpdateMode.OnPropertyChanged
                                )
                            );
                    }

                    ctrl.DataBindings.Clear();

                    foreach (var newBinding in newBindingList)
                    {
                        ctrl.DataBindings.Add(newBinding);
                    }
                }

                if (ctrl is CommandItemControl)
                {
                    var control = (CommandItemControl)ctrl;
                    switch (control.HelperType)
                    {
                        case UIAdditionalHelperType.ShowVariableHelper:
                            control.DataSource = editor.ScriptContext.Variables;
                            break;
                        case UIAdditionalHelperType.ShowElementHelper:
                            control.DataSource = editor.ScriptContext.Elements;
                            break;
                        default:
                            break;
                    }
                }
            }
        }
    }
}
