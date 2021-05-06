﻿using Microsoft.Office.Interop.Word;
using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Properties;
using OpenBots.Core.Utilities.CommonUtilities;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Word.Application;
using DataTable = System.Data.DataTable;
using Tasks = System.Threading.Tasks;

namespace OpenBots.Commands.Word
{
	[Serializable]
	[Category("Word Commands")]
	[Description("This command appends a DataTable to a Word Document.")]
	public class WordAppendDataTableCommand : ScriptCommand
	{
		[Required]
		[DisplayName("Word Instance Name")]
		[Description("Enter the unique instance that was specified in the **Create Application** command.")]
		[SampleUsage("MyWordInstance")]
		[Remarks("Failure to enter the correct instance or failure to first call the **Create Application** command will cause an error.")]
		[CompatibleTypes(new Type[] { typeof(Application) })]
		public string v_InstanceName { get; set; }

		[Required]
		[DisplayName("DataTable")]
		[Description("Enter the DataTable to append to the Document.")]
		[SampleUsage("vDataTable")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(DataTable) })]
		public string v_DataTable { get; set; }

		public WordAppendDataTableCommand()
		{
			CommandName = "WordAppendDataTableCommand";
			SelectionName = "Append DataTable";
			CommandEnabled = true;
			CommandIcon = Resources.command_files;

			v_InstanceName = "DefaultWord";
		}
		public async override Tasks.Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var wordObject = v_InstanceName.GetAppInstance(engine);

			DataTable dataTable = (DataTable)await v_DataTable.EvaluateCode(engine);

			//selecting the word instance and open document
			Application wordInstance = (Application)wordObject;
			Document wordDocument = wordInstance.ActiveDocument;

			//converting System DataTable to Word DataTable
			int RowCount = dataTable.Rows.Count; 
			int ColumnCount = dataTable.Columns.Count;
			object[,] DataArray = new object[RowCount, ColumnCount];
		   
			int r = 0;
			for (int c = 0; c < ColumnCount; c++)
			{
				for (r = 0; r < RowCount; r++)
				{
					DataArray[r, c] = dataTable.Rows[r][c];
				} //end row loop
			} //end column loop

			object collapseEnd = WdCollapseDirection.wdCollapseEnd;
			dynamic docRange = wordDocument.Content; 
			docRange.Collapse(ref collapseEnd);

			string tempString = "";
			for (r = 0; r <= RowCount - 1; r++)
			{
				for (int c = 0; c <= ColumnCount - 1; c++)
					tempString = tempString + DataArray[r, c] + "\t";
			}

			//appending data row text after all text/images
			docRange.Text = tempString;

			//converting and formatting data table
			object Separator = WdTableFieldSeparator.wdSeparateByTabs;
			object Format = WdTableFormat.wdTableFormatGrid1;
			object ApplyBorders = true;
			object AutoFit = true;
			object AutoFitBehavior = WdAutoFitBehavior.wdAutoFitContent;
			docRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount, Type.Missing, ref Format,
									ref ApplyBorders, Type.Missing, Type.Missing, Type.Missing,Type.Missing, 
									Type.Missing, Type.Missing, Type.Missing, ref AutoFit, ref AutoFitBehavior, 
									Type.Missing);

			docRange.Select();
			wordDocument.Application.Selection.Tables[1].Select();
			wordDocument.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
			wordDocument.Application.Selection.Tables[1].Rows.Alignment = 0;
			wordDocument.Application.Selection.Tables[1].Rows[1].Select();
			wordDocument.Application.Selection.InsertRowsAbove(1);
			wordDocument.Application.Selection.Tables[1].Rows[1].Select();

			//Adding header row manually
			for (int c = 0; c <= ColumnCount - 1; c++)
				wordDocument.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = dataTable.Columns[c].ColumnName;

			//Formatting header row
			wordDocument.Application.Selection.Tables[1].Rows[1].Select();
			wordDocument.Application.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
			wordDocument.Application.Selection.Font.Bold = 1;

			int docRowCount = wordDocument.Application.Selection.Tables[1].Rows.Count;
			wordDocument.Application.Selection.Tables[1].Rows[docRowCount].Delete();
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_DataTable", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Append '{v_DataTable}' - Instance Name '{v_InstanceName}']";
		}       
	}
}