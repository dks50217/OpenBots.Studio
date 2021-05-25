﻿using Microsoft.Office.Interop.Excel;
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
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace OpenBots.Commands.Excel
{
	[Serializable]
	[Category("Excel Commands")]
	[Description("This command gets text from a specific cell in an Excel Worksheet.")]
	public class ExcelGetCellCommand : ScriptCommand
	{
		[Required]
		[DisplayName("Excel Instance Name")]
		[Description("Enter the unique instance that was specified in the **Create Application** command.")]
		[SampleUsage("MyExcelInstance")]
		[Remarks("Failure to enter the correct instance or failure to first call the **Create Application** command will cause an error.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(OBAppInstance) })]
		public string v_InstanceName { get; set; }

		[Required]
		[DisplayName("Cell Location")]
		[Description("Enter the location of the cell to extract.")]
		[SampleUsage("\"A1\" || vCellLocation")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_CellLocation { get; set; }

		[Required]
		[DisplayName("Read Formulas")]
		[PropertyUISelectionOption("Yes")]
		[PropertyUISelectionOption("No")]
		[Description("When selected, formulas will be extracted rather than their calculated values.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_Formulas { get; set; }

		[Required]
		[Editable(false)]
		[DisplayName("Output Cell Value Variable")]
		[Description("Create a new variable or select a variable from the list.")]
		[SampleUsage("vUserVariable")]
		[Remarks("New variables/arguments may be instantiated by utilizing the Ctrl+K/Ctrl+J shortcuts.")]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_OutputUserVariableName { get; set; }

		public ExcelGetCellCommand()
		{
			CommandName = "ExcelGetCellCommand";
			SelectionName = "Get Cell";
			CommandEnabled = true;
			CommandIcon = Resources.command_excel;

			v_InstanceName = "DefaultExcel";
			v_Formulas = "No";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var excelObject = ((OBAppInstance)await v_InstanceName.EvaluateCode(engine)).Value;
			var vTargetAddress = (string)await v_CellLocation.EvaluateCode(engine);
			var excelInstance = (Application)excelObject;
			Worksheet excelSheet = excelInstance.ActiveSheet;
            string cellValue;

            if (v_Formulas == "Yes")
				cellValue = (string)excelSheet.Range[vTargetAddress].Formula;
			else
				cellValue = (string)excelSheet.Range[vTargetAddress].Value.ToString();

			cellValue.SetVariableValue(engine, v_OutputUserVariableName);          
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_CellLocation", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_Formulas", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultOutputGroupFor("v_OutputUserVariableName", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Get Value From '{v_CellLocation}' - Store Cell Value in '{v_OutputUserVariableName}' - Instance Name '{v_InstanceName}']";
		}
	}
}