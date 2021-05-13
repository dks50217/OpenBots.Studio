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
	[Description("This command activates a specific Worksheet in an Excel Workbook.")]
	public class ExcelActivateSheetCommand : ScriptCommand
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
		[DisplayName("Worksheet Name")]
		[Description("Specify the Worksheet within the Workbook to activate.")]
		[SampleUsage("\"Sheet1\" || vSheet")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_SheetName { get; set; }

		public ExcelActivateSheetCommand()
		{
			CommandName = "ExcelActivateSheetCommand";
			SelectionName = "Activate Sheet";
			CommandEnabled = true;
			CommandIcon = Resources.command_excel;

			v_InstanceName = "DefaultExcel";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			string vSheetToActivate = (string)await v_SheetName.EvaluateCode(engine);

			var excelObject = ((OBAppInstance)await v_InstanceName.EvaluateCode(engine)).Value;
			var excelInstance = (Application)excelObject;      
			var workSheet = excelInstance.Sheets[vSheetToActivate] as Worksheet;
			workSheet.Select(); 
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_SheetName", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Sheet '{v_SheetName}' - Instance Name '{v_InstanceName}']";
		}
	}
}