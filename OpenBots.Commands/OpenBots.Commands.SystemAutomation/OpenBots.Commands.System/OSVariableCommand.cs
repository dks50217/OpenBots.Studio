﻿using Newtonsoft.Json;
using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Properties;
using OpenBots.Core.Utilities.CommonUtilities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.Management;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace OpenBots.Commands.System
{
	[Serializable]
	[Category("System Commands")]
	[Description("This command exclusively selects an OS variable.")]
	public class OSVariableCommand : ScriptCommand
	{

		[Required]
		[DisplayName("OS Variable")]
		[Description("Select an OS variable from one of the options.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_OSVariableName { get; set; }

		[Required]
		[Editable(false)]
		[DisplayName("Output OS Variable")]
		[Description("Create a new variable or select a variable from the list.")]
		[SampleUsage("vUserVariable")]
		[Remarks("New variables/arguments may be instantiated by utilizing the Ctrl+K/Ctrl+J shortcuts.")]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_OutputUserVariableName { get; set; }

		[JsonIgnore]
		[Browsable(false)]
		private ComboBox _variableNameComboBox;

		[JsonIgnore]
		[Browsable(false)]
		private Label _variableValue;

		public OSVariableCommand()
		{
			CommandName = "OSVariableCommand";
			SelectionName = "OS Variable";
			CommandEnabled = true;
			CommandIcon = Resources.command_system;

		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var systemVariable = (string)await v_OSVariableName.EvaluateCode(engine);

			ObjectQuery wql = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
			ManagementObjectSearcher searcher = new ManagementObjectSearcher(wql);
			ManagementObjectCollection results = searcher.Get();

			foreach (ManagementObject result in results)
			{
				foreach (PropertyData prop in result.Properties)
				{
					if (prop.Name == systemVariable.ToString())
					{
						var sysValue = prop.Value.ToString();
						sysValue.SetVariableValue(engine, v_OutputUserVariableName);
						return;
					}
				}
			}
			throw new Exception("System Property '" + systemVariable + "' not found!");
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			var ActionNameComboBoxLabel = commandControls.CreateDefaultLabelFor("v_OSVariableName", this);
			_variableNameComboBox = commandControls.CreateDropdownFor("v_OSVariableName", this);

			ObjectQuery wql = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
			ManagementObjectSearcher searcher = new ManagementObjectSearcher(wql);
			ManagementObjectCollection results = searcher.Get();

			foreach (ManagementObject result in results)
			{
				foreach (PropertyData prop in result.Properties)
					_variableNameComboBox.Items.Add($"\"{prop.Name}\"");
			}

			_variableNameComboBox.SelectedValueChanged += VariableNameComboBox_SelectedValueChanged;
			RenderedControls.Add(ActionNameComboBoxLabel);
			RenderedControls.Add(_variableNameComboBox);

			_variableValue = new Label();
			_variableValue.Font = new Font("Segoe UI Semilight", 10, FontStyle.Bold);
			_variableValue.ForeColor = Color.White;
			RenderedControls.Add(_variableValue);

			RenderedControls.AddRange(commandControls.CreateDefaultOutputGroupFor("v_OutputUserVariableName", this, editor));      

			return RenderedControls;
		}

		private void VariableNameComboBox_SelectedValueChanged(object sender, EventArgs e)
		{
			string selectedValue;

			if (_variableNameComboBox.SelectedItem == null)
				return;
			else
				selectedValue = _variableNameComboBox.SelectedItem.ToString().Trim('\"');

			ObjectQuery wql = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
			ManagementObjectSearcher searcher = new ManagementObjectSearcher(wql);
			ManagementObjectCollection results = searcher.Get();

			foreach (ManagementObject result in results)
			{
				foreach (PropertyData prop in result.Properties)
				{
					if (prop.Name == selectedValue)
					{
						_variableValue.Text = prop.Value?.ToString();
						return;
					}
				}
			}
			_variableValue.Text = "[ex. **Item not found**]";
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Store OS Variable '{v_OSVariableName}' in '{v_OutputUserVariableName}']";
		}
	}


}