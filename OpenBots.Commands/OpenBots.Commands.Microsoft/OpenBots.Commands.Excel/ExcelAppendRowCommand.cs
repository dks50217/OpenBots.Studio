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
using System.Data;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace OpenBots.Commands.Excel
{
	[Serializable]
	[Category("Excel Commands")]
	[Description("This command appends a row after the last row of an Excel Worksheet.")]
	public class ExcelAppendRowCommand : ScriptCommand
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
		[DisplayName("Row")]
		[Description("Enter the row value to append.")]
		[SampleUsage("new List<string>() { \"Hello\", \"World\" } || vList || vDataRow")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(DataRow), typeof(List<string>) })]
		public string v_RowToSet { get; set; }

		public ExcelAppendRowCommand()
		{
			CommandName = "ExcelAppendRowCommand";
			SelectionName = "Append Row";
			CommandEnabled = true;
			CommandIcon = Resources.command_excel;

			v_InstanceName = "DefaultExcel";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			dynamic vRow = await v_RowToSet.EvaluateCode(engine);

			var excelObject = ((OBAppInstance)await v_InstanceName.EvaluateCode(engine)).Value;
			var excelInstance = (Application)excelObject;
			Worksheet excelSheet = excelInstance.ActiveSheet;

			int lastUsedRow;
			int i = 1;
			try
			{
				lastUsedRow = excelSheet.Cells.Find("*", Missing.Value, Missing.Value, Missing.Value, XlSearchOrder.xlByRows, 
													XlSearchDirection.xlPrevious, false, Missing.Value, Missing.Value).Row;
			}
			catch(Exception)
			{
				lastUsedRow = 0;
			}

			DataRow row;
			if (vRow != null && vRow is DataRow)
			{
				row = (DataRow)vRow;

				string cellValue;
				for (int j = 0; j < row.ItemArray.Length; j++)
				{
					if (row.ItemArray[j] == null)
						cellValue = string.Empty;
					else
						cellValue = row.ItemArray[j].ToString();

					excelSheet.Cells[lastUsedRow + 1, i] = cellValue;
					i++;
				}
			}
			else
			{
				var vRowList = (List<string>)vRow;
				string cellValue;
				foreach (var item in vRowList)
				{
					cellValue = item;
					if (cellValue == "null")
					{
						cellValue = string.Empty;
					}
					excelSheet.Cells[lastUsedRow + 1, i] = cellValue;
					i++;
				}
			}          
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_RowToSet", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Append '{v_RowToSet}' - Instance Name '{v_InstanceName}']";
		}        
	}
}
