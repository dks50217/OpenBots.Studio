﻿using OpenBots.Core.Attributes.PropertyAttributes;
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
	[Description("This command runs a macro in an Excel Workbook.")]

	public class ExcelRunMacroCommand : ScriptCommand
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
		[DisplayName("Macro Name")]
		[Description("Enter the name of the macro as it exists in the Worksheet.")]
		[SampleUsage("\"Macro1\" || vMacro")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_MacroName { get; set; }

		public ExcelRunMacroCommand()
		{
			CommandName = "ExcelRunMacroCommand";
			SelectionName = "Run Macro";
			CommandEnabled = true;
			CommandIcon = Resources.command_excel;

			v_InstanceName = "DefaultExcel";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var vMacro = (string)await v_MacroName.EvaluateCode(engine);
			var excelObject = ((OBAppInstance)await v_InstanceName.EvaluateCode(engine)).Value;
			var excelInstance = (Application)excelObject;
			excelInstance.Run(vMacro);           
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_MacroName", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Run Macro '{v_MacroName}' - Instance Name '{v_InstanceName}']";
		}
	}
}