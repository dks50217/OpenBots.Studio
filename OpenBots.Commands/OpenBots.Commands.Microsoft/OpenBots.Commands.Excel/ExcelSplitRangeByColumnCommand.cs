﻿using Microsoft.Office.Interop.Excel;
using OpenBots.Commands.Microsoft.Library;
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
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using DataTable = System.Data.DataTable;

namespace OpenBots.Commands.Excel
{
    [Serializable]
	[Category("Excel Commands")]
	[Description("This command takes a specific Excel range, splits it into separate ranges by column, and stores them in new Workbooks.")]
	public class ExcelSplitRangeByColumnCommand : ScriptCommand
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
		[DisplayName("Range")]
		[Description("Enter the location of the range to split.")]
		[SampleUsage("\"A1:B10\" || \"A1:\" || vRange || vStart + \":\" + vEnd || vStart + \":\"")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_Range { get; set; }

		[Required]
		[DisplayName("Column to Split")]
		[Description("Enter the name of the column you wish to split the selected range by.")]
		[SampleUsage("\"ColA\" || vColumnName")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_ColumnName { get; set; }

		[Required]
		[DisplayName("Split Range Output Directory")]
		[Description("Enter or Select the new directory for the split range files.")]
		[SampleUsage("@\"C:\\temp\" || ProjectPath + @\"\\temp\" || vDirectoryPath")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[Editor("ShowFolderSelectionHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_OutputDirectory { get; set; }

		[Required]
		[DisplayName("Output File Type")]
		[PropertyUISelectionOption("xlsx")]
		[PropertyUISelectionOption("csv")]
		[Description("Specify the file format type for the split range files.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_FileType { get; set; }

		[Required]
		[Editable(false)]
		[DisplayName("Output DataTable List Variable")]
		[Description("Create a new variable or select a variable from the list.")]
		[SampleUsage("vUserVariable")]
		[Remarks("New variables/arguments may be instantiated by utilizing the Ctrl+K/Ctrl+J shortcuts.")]
		[CompatibleTypes(new Type[] { typeof(List<DataTable>) })]
		public string v_OutputUserVariableName { get; set; }

		public ExcelSplitRangeByColumnCommand()
		{
			CommandName = "ExcelSplitRangeByColumnCommand";
			SelectionName = "Split Range By Column";
			CommandEnabled = true;
			CommandIcon = Resources.command_excel;

			v_FileType = "xlsx";
			v_Range = "\"A1:\"";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var vExcelObject = ((OBAppInstance)await v_InstanceName.EvaluateCode(engine)).Value;
			var vRange = (string)await v_Range.EvaluateCode(engine);
			var vColumnName = (string)await v_ColumnName.EvaluateCode(engine);
			var vOutputDirectory = (string)await v_OutputDirectory.EvaluateCode(engine);
			var excelInstance = (Application)vExcelObject;

			excelInstance.DisplayAlerts = false;
			Worksheet excelSheet = excelInstance.ActiveSheet;
			Range cellRange = excelInstance.GetRange(vRange, excelSheet);

			//Convert Range to DataTable
			List<object> lst = new List<object>();
			int rw = cellRange.Rows.Count;
			int cl = cellRange.Columns.Count;
			int rCnt;
			int cCnt;
			string cName;
			DataTable DT = new DataTable();
			
			//start from row 2
			for (rCnt = 2; rCnt <= rw; rCnt++)
			{
				DataRow newRow = DT.NewRow();
				for (cCnt = 1; cCnt <= cl; cCnt++)
				{
					if (((cellRange.Cells[rCnt, cCnt] as Range).Value2) != null)
					{
						if (!DT.Columns.Contains(cCnt.ToString()))
						{
							DT.Columns.Add(cCnt.ToString());
						}
						newRow[cCnt.ToString()] = ((cellRange.Cells[rCnt, cCnt] as Range).Value2).ToString();
					}
					else if (((cellRange.Cells[rCnt, cCnt] as Range).Value2) == null && ((cellRange.Cells[1, cCnt] as Range).Value2) != null)
					{
						if (!DT.Columns.Contains(cCnt.ToString()))
						{
							DT.Columns.Add(cCnt.ToString());
						}
						newRow[cCnt.ToString()] = string.Empty;
					}
				}
				DT.Rows.Add(newRow);
			}

			//Set column names
			for (cCnt = 1; cCnt <= cl; cCnt++)
			{
				cName = ((cellRange.Cells[1, cCnt] as Range).Value2).ToString();
				DT.Columns[cCnt-1].ColumnName = cName;
			}

			//split table by column
			List<DataTable> result = DT.AsEnumerable()
									   .GroupBy(row => row.Field<string>(vColumnName))
									   .Select(g => g.CopyToDataTable())
									   .ToList();

			//add list of datatables to output variable
			result.SetVariableValue(engine, v_OutputUserVariableName);

			//save split datatables in individual workbooks labeled by selected column data
			if (Directory.Exists(vOutputDirectory))
			{
				string newName;
				foreach (DataTable newDT in result)
				{
					try
					{
						newName = newDT.Rows[0].Field<string>(vColumnName).ToString();
					}
					catch (Exception)
					{
						continue;
					}
					var splitExcelInstance = new Application();
					splitExcelInstance.DisplayAlerts = false;

					Workbook newWorkBook = splitExcelInstance.Workbooks.Add(Type.Missing);
					Worksheet newSheet = newWorkBook.ActiveSheet;
					for (int i = 1; i < newDT.Columns.Count + 1; i++)
					{
						newSheet.Cells[1, i] = newDT.Columns[i - 1].ColumnName;
					}

					for (int j = 0; j < newDT.Rows.Count; j++)
					{
						for (int k = 0; k < newDT.Columns.Count; k++)
						{
							newSheet.Cells[j + 2, k + 1] = newDT.Rows[j].ItemArray[k].ToString();
						}
					}

					if (string.IsNullOrEmpty(newName))
						newName = "_";

					newName = string.Join("_", newName.Split(Path.GetInvalidFileNameChars()));

					if (v_FileType == "csv" && !newName.Equals(string.Empty))
					{
						newWorkBook.SaveAs(Path.Combine(vOutputDirectory, newName), XlFileFormat.xlCSV, Type.Missing, Type.Missing,
										Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange,
										Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);                      
					}
					else if (!newName.Equals(string.Empty))
					{
						newWorkBook.SaveAs(Path.Combine(vOutputDirectory, newName + ".xlsx"));
					}
					newWorkBook.Close();
				}
			}   
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_Range", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_ColumnName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_OutputDirectory", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_FileType", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultOutputGroupFor("v_OutputUserVariableName", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Split Range '{v_Range}' by Column '{v_ColumnName}' - Store DataTable List in '{v_OutputUserVariableName}' - Instance Name '{v_InstanceName}']";
		}
	}
}
