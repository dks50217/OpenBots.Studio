﻿using Newtonsoft.Json;
using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Model.ApplicationModel;
using OpenBots.Core.Properties;
using OpenBots.Core.Utilities.CommonUtilities;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace OpenBots.Commands.Excel
{
	[Serializable]
	[Category("Excel Commands")]
	[Description("This command creates an Excel Instance.")]
	public class ExcelCreateApplicationCommand : ScriptCommand
	{
		[Required]
		[DisplayName("Excel Instance Name")]
		[Description("Enter a unique name that will represent the application instance.")]
		[SampleUsage("MyExcelInstance")]
		[Remarks("This unique name allows you to refer to the instance by name in future commands, " +
				 "ensuring that the commands you specify run against the correct application.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(OBAppInstance) })]
		public string v_InstanceName { get; set; }

		[Required]
		[DisplayName("New/Open Workbook")]
		[PropertyUISelectionOption("New Workbook")]
		[PropertyUISelectionOption("Open Workbook")]
		[Description("Indicate whether to create a new Workbook or to open an existing Workbook.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_NewOpenWorkbook { get; set; }

		[Required]
		[DisplayName("Workbook File Path")]
		[Description("Enter or Select the path to the Workbook file.")]
		[SampleUsage("@\"C:\\temp\\myfile.xlsx\" || ProjectPath + @\"\\myfile.xlsx\" || vFilePath")]
		[Remarks("This input should only be used for opening existing Workbooks.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[Editor("ShowFileSelectionHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_FilePath { get; set; }

		[Required]
		[DisplayName("Visible")]
		[PropertyUISelectionOption("Yes")]
		[PropertyUISelectionOption("No")]
		[Description("Indicate whether the Excel automation should be visible or not.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_Visible { get; set; }

		[Required]
		[DisplayName("Close All Existing Excel Instances")]
		[PropertyUISelectionOption("Yes")]
		[PropertyUISelectionOption("No")]
		[Description("Indicate whether to close any existing Excel instances before executing Excel Automation.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_CloseAllInstances { get; set; }

		[JsonIgnore]
		[Browsable(false)]
		private List<Control> _openFileControls;

		[JsonIgnore]
		[Browsable(false)]
		private bool _hasRendered;

		public ExcelCreateApplicationCommand()
		{
			CommandName = "ExcelCreateApplicationCommand";
			SelectionName = "Create Excel Application";
			CommandEnabled = true;
			CommandIcon = Resources.command_spreadsheet;

			v_InstanceName = "DefaultExcel";
			v_NewOpenWorkbook = "New Workbook";
			v_Visible = "No";
			v_CloseAllInstances = "Yes";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var vFilePath = (string)await v_FilePath.EvaluateCode(engine);

			if (v_CloseAllInstances == "Yes")
			{
				var processes = Process.GetProcessesByName("excel");
				foreach (var prc in processes)
				{
					prc.Kill();
				}
			}

			var newExcelSession = new Application();
			if (v_Visible == "Yes")
				newExcelSession.Visible = true;
			else
				newExcelSession.Visible = false;

			new OBAppInstance(v_InstanceName, newExcelSession).SetVariableValue(engine, v_InstanceName);

			if (v_NewOpenWorkbook == "New Workbook")
			{
				if (!string.IsNullOrEmpty(vFilePath))
					throw new InvalidDataException("File path should not be provided for a new Excel Workbook");
				else
					newExcelSession.Workbooks.Add();
			}
			else if (v_NewOpenWorkbook == "Open Workbook")
			{
				if (string.IsNullOrEmpty(vFilePath))
					throw new NullReferenceException("File path for Excel Workbook not provided");
				else
					newExcelSession.Workbooks.Open(vFilePath);
			}
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_NewOpenWorkbook", this, editor));
			((ComboBox)RenderedControls[3]).SelectedIndexChanged += OpenFileComboBox_SelectedIndexChanged;

			_openFileControls = new List<Control>();
			_openFileControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_FilePath", this, editor));

			RenderedControls.AddRange(_openFileControls);
			
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_Visible", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_CloseAllInstances", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [{v_NewOpenWorkbook} - Visible '{v_Visible}' - Close Instances '{v_CloseAllInstances}' - New Instance Name '{v_InstanceName}']";
		}

		public override void Shown()
		{
			base.Shown();
			_hasRendered = true;
			OpenFileComboBox_SelectedIndexChanged(this, null);
		}

		private void OpenFileComboBox_SelectedIndexChanged(object sender, EventArgs e)
		{
			if (((ComboBox)RenderedControls[3]).Text == "Open Workbook" && _hasRendered)
			{
				foreach (var ctrl in _openFileControls)
					ctrl.Visible = true;
			}
			else if(_hasRendered)
			{
				foreach (var ctrl in _openFileControls)
				{
					ctrl.Visible = false;
					if (ctrl is TextBox)
						((TextBox)ctrl).Clear();
				}
			}
		}
	}
}