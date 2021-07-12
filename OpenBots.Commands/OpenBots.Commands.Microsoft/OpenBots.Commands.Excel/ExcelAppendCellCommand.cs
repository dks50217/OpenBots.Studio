﻿using Microsoft.Office.Interop.Excel;
using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Interfaces;
using OpenBots.Core.Model.ApplicationModel;
using OpenBots.Core.Properties;
using OpenBots.Core.Utilities.CommonUtilities;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace OpenBots.Commands.Excel
{
	[Serializable]
	[Category("Excel Commands")]
	[Description("This command appends a cell to the first column after the last row in an Excel Worksheet.")]
	public class ExcelAppendCellCommand : ScriptCommand
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
		[DisplayName("Cell Value")]
		[Description("Enter the text value that will be set in the appended cell.")]
		[SampleUsage("\"Hello World\" || vText")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_TextToSet { get; set; }

		public ExcelAppendCellCommand()
		{
			CommandName = "ExcelAppendCellCommand";
			SelectionName = "Append Cell";
			CommandEnabled = true;
			CommandIcon = Resources.command_excel;
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var excelObject = ((OBAppInstance)await v_InstanceName.EvaluateCode(engine)).Value;
			var excelInstance = (Application)excelObject;
			Worksheet excelSheet = excelInstance.ActiveSheet;
  
			var lastUsedRow = excelSheet.Cells.Find("*", Missing.Value, Missing.Value, Missing.Value, XlSearchOrder.xlByRows, 
													XlSearchDirection.xlPrevious, false, Missing.Value, Missing.Value).Row;
			var targetAddress = "A" + (lastUsedRow + 1);
			var vTargetText = (string)await v_TextToSet.EvaluateCode(engine);
			excelSheet.Range[targetAddress].Value = vTargetText;
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_TextToSet", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Append '{v_TextToSet}' - Instance Name '{v_InstanceName}']";
		}
	}
}