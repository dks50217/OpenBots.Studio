﻿using Newtonsoft.Json;
using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Properties;
using OpenBots.Core.Script;
using OpenBots.Core.UI.Controls;
using OpenBots.Core.Utilities.CommandUtilities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Tasks = System.Threading.Tasks;

namespace OpenBots.Commands.If
{
	[Serializable]
	[Category("If Commands")]
	[Description("This command evaluates a group of combined logical statements to determine if the combined result of the statements is 'true' or 'false' and subsequently performs action(s) based on the result.")]
	public class BeginMultiIfCommand : ScriptCommand
	{
		[Required]
		[DisplayName("Logic Type")]
		[PropertyUISelectionOption("And")]
		[PropertyUISelectionOption("Or")]
		[Description("Select the logic to use when evaluating multiple Ifs.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_LogicType { get; set; }

		[Required]
		[DisplayName("Multiple If Conditions")]
		[Description("Add new If condition(s).")]
		[SampleUsage("")]
		[Remarks("")]
		[Editor("ShowIfBuilder", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(Bitmap), typeof(DateTime), typeof(string), typeof(double), typeof(int), typeof(bool) })]
		public DataTable v_IfConditionsTable { get; set; }

		[JsonIgnore]
		[Browsable(false)]
		private DataGridView _ifConditionHelper;

		public BeginMultiIfCommand()
		{
			CommandName = "BeginMultiIfCommand";
			SelectionName = "Begin Multi If";
			CommandEnabled = true;
			CommandIcon = Resources.command_begin_multi_if;
			ScopeStartCommand = true;

			v_LogicType = "And";
			v_IfConditionsTable = new DataTable();
			v_IfConditionsTable.TableName = DateTime.Now.ToString("MultiIfConditionTable" + DateTime.Now.ToString("MMddyy.hhmmss"));
			v_IfConditionsTable.Columns.Add("Statement");
			v_IfConditionsTable.Columns.Add("CommandData");
		}
	   
		public async override Tasks.Task RunCommand(object sender, ScriptAction parentCommand)
		{
			var engine = (IAutomationEngineInstance)sender;

			bool isTrueStatement = true;
			foreach (DataRow rw in v_IfConditionsTable.Rows)
			{
				var commandData = rw["CommandData"].ToString();
				var ifCommand = JsonConvert.DeserializeObject<BeginIfCommand>(commandData);
				var statementResult = await CommandsHelper.DetermineStatementTruth(engine, ifCommand.v_IfActionType, ifCommand.v_ActionParameterTable);

				if (!statementResult && v_LogicType == "And")
				{
					isTrueStatement = false;
					break;
				}

				if(statementResult && v_LogicType == "Or")
                {
					isTrueStatement = true;
					break;
                }
				else if (v_LogicType == "Or")
                {
					isTrueStatement = false;
                }
			}

			//report evaluation
			if (isTrueStatement)
			{
				engine.ReportProgress("If Conditions Evaluated True");
			}
			else
			{
				engine.ReportProgress("If Conditions Evaluated False");
			}
			
			int startIndex, endIndex, elseIndex;
			if (parentCommand.AdditionalScriptCommands.Any(item => item.ScriptCommand is ElseCommand))
			{
				elseIndex = parentCommand.AdditionalScriptCommands.FindIndex(a => a.ScriptCommand is ElseCommand);

				if (isTrueStatement)
				{
					startIndex = 0;
					endIndex = elseIndex;
				}
				else
				{
					startIndex = elseIndex + 1;
					endIndex = parentCommand.AdditionalScriptCommands.Count;
				}
			}
			else if (isTrueStatement)
			{
				startIndex = 0;
				endIndex = parentCommand.AdditionalScriptCommands.Count;
			}
			else
			{
				return;
			}

			for (int i = startIndex; i < endIndex; i++)
			{
				if ((engine.IsCancellationPending) || (engine.CurrentLoopCancelled))
					return;

				await engine.ExecuteCommand(parentCommand.AdditionalScriptCommands[i]);
			}
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_LogicType", this, editor));

			//create controls
			var controls = commandControls.CreateDefaultDataGridViewGroupFor("v_IfConditionsTable", this, editor);
			_ifConditionHelper = controls[2] as DataGridView;

			//handle helper click
			var helper = controls[1] as CommandItemControl;
			helper.Click += (sender, e) => CreateIfCondition(sender, e, editor, commandControls);

			//add for rendering
			RenderedControls.AddRange(controls);

			//define if condition helper
			_ifConditionHelper.Width = 450;
			_ifConditionHelper.Height = 200;
			_ifConditionHelper.AutoGenerateColumns = false;
			_ifConditionHelper.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
			_ifConditionHelper.Columns.Add(new DataGridViewTextBoxColumn() { HeaderText = "Condition", DataPropertyName = "Statement", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
			_ifConditionHelper.Columns.Add(new DataGridViewTextBoxColumn() { HeaderText = "CommandData", DataPropertyName = "CommandData", ReadOnly = true, Visible = false });
			_ifConditionHelper.Columns.Add(new DataGridViewButtonColumn() { HeaderText = "Edit", UseColumnTextForButtonValue = true,  Text = "Edit", Width = 45 });
			_ifConditionHelper.Columns.Add(new DataGridViewButtonColumn() { HeaderText = "Delete", UseColumnTextForButtonValue = true, Text = "Delete", Width = 60 });
			_ifConditionHelper.AllowUserToAddRows = false;
			_ifConditionHelper.AllowUserToDeleteRows = true;
			_ifConditionHelper.CellContentClick += (sender, e) => IfConditionHelper_CellContentClick(sender, e, editor, commandControls);

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			if (v_IfConditionsTable.Rows.Count == 0)
			{
				return "If <Not Configured>";
			}
			else if (v_LogicType == "And")
			{
				var statements = v_IfConditionsTable.AsEnumerable().Select(f => f.Field<string>("Statement")).ToList();
				return string.Join(" && ", statements);
			}
			else
            {
				var statements = v_IfConditionsTable.AsEnumerable().Select(f => f.Field<string>("Statement")).ToList();
				return string.Join(" || ", statements);
			}
		}

		private void IfConditionHelper_CellContentClick(object sender, DataGridViewCellEventArgs e, IfrmCommandEditor parentEditor, ICommandControls commandControls)
		{
			var senderGrid = (DataGridView)sender;

			if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex >= 0)
			{
				var buttonSelected = senderGrid.Rows[e.RowIndex].Cells[e.ColumnIndex] as DataGridViewButtonCell;
				var selectedRow = v_IfConditionsTable.Rows[e.RowIndex];

				if (buttonSelected.Value.ToString() == "Edit")
				{
					//launch editor
					var statement = selectedRow["Statement"];
					var commandData = selectedRow["CommandData"].ToString();

					var ifCommand = JsonConvert.DeserializeObject<BeginIfCommand>(commandData);

					var automationCommands = new List<AutomationCommand>() { CommandsHelper.ConvertToAutomationCommand(typeof(BeginIfCommand)) };
					IfrmCommandEditor editor = commandControls.CreateCommandEditorForm(automationCommands, null);
					editor.SelectedCommand = ifCommand;
					editor.EditingCommand = ifCommand;
					editor.OriginalCommand = ifCommand;
					editor.CreationModeInstance = CreationMode.Edit;
					editor.ScriptContext = parentEditor.ScriptContext;
					editor.TypeContext = parentEditor.TypeContext;

					if (((Form)editor).ShowDialog() == DialogResult.OK)
					{
						var editedCommand = editor.SelectedCommand as BeginIfCommand;
						var displayText = editedCommand.GetDisplayValue();
						var serializedData = JsonConvert.SerializeObject(editedCommand);

						selectedRow["Statement"] = displayText;
						selectedRow["CommandData"] = serializedData;
					}
				}
				else if (buttonSelected.Value.ToString() == "Delete")
				{
					//delete
					v_IfConditionsTable.Rows.Remove(selectedRow);
				}
				else
				{
					throw new NotImplementedException("Requested Action is not implemented.");
				}
			}
		}

		private void CreateIfCondition(object sender, EventArgs e, IfrmCommandEditor parentEditor, ICommandControls commandControls)
		{
			var automationCommands = new List<AutomationCommand>() { CommandsHelper.ConvertToAutomationCommand(typeof(BeginIfCommand)) };
			IfrmCommandEditor editor = commandControls.CreateCommandEditorForm(automationCommands, null);
            editor.SelectedCommand = new BeginIfCommand();
			editor.ScriptContext = parentEditor.ScriptContext;
			editor.TypeContext = parentEditor.TypeContext;

			if (((Form)editor).ShowDialog() == DialogResult.OK)
			{
				//get data
				var configuredCommand = editor.SelectedCommand as BeginIfCommand;
				var displayText = configuredCommand.GetDisplayValue();
				var serializedData = JsonConvert.SerializeObject(configuredCommand);
				parentEditor.ScriptContext = editor.ScriptContext;
				parentEditor.TypeContext = editor.TypeContext;

				//add to list
				v_IfConditionsTable.Rows.Add(displayText, serializedData);
			}
		}
	}
}