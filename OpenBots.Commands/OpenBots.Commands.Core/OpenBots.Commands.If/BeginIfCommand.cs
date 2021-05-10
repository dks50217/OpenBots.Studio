﻿using Newtonsoft.Json;
using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Model.ApplicationModel;
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
	[Description("This command evaluates a logical statement to determine if the statement is 'true' or 'false' and subsequently performs action(s) based on either condition.")]
	public class BeginIfCommand : ScriptCommand, IConditionCommand
	{
		[Required]
		[DisplayName("Condition Type")]
		[PropertyUISelectionOption("Number Compare")]
		[PropertyUISelectionOption("Date Compare")]
		[PropertyUISelectionOption("Text Compare")]
		[PropertyUISelectionOption("Has Value")]
		[PropertyUISelectionOption("Is Numeric")]
		[PropertyUISelectionOption("Window Name Exists")]
		[PropertyUISelectionOption("Active Window Name Is")]
		[PropertyUISelectionOption("File Exists")]
		[PropertyUISelectionOption("Folder Exists")]
		[PropertyUISelectionOption("Web Element Exists")]
		[PropertyUISelectionOption("GUI Element Exists")]
		[PropertyUISelectionOption("Image Element Exists")]
		[PropertyUISelectionOption("App Instance Exists")]
		[PropertyUISelectionOption("Error Occured")]
		[PropertyUISelectionOption("Error Did Not Occur")]
		[Description("Select the necessary condition type.")]
		[Remarks("")]
		public string v_IfActionType { get; set; }

		[Required]
		[DisplayName("Additional Parameters")]
		[Description("Supply or Select the required comparison parameters.")]
		[SampleUsage("Param Value || vParamValue")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(Bitmap), typeof(DateTime), typeof(string), typeof(double), typeof(int), typeof(bool), typeof(OBAppInstance) })]
		public DataTable v_ActionParameterTable { get; set; }

		[JsonIgnore]
		[Browsable(false)]
		private DataGridView _ifGridViewHelper;

		[JsonIgnore]
		[Browsable(false)]
		private ComboBox _actionDropdown;

		[JsonIgnore]
		[Browsable(false)]
		private List<Control> _parameterControls;

		[JsonIgnore]
		[Browsable(false)]
		private CommandItemControl _recorderControl;

		public BeginIfCommand()
		{
			CommandName = "BeginIfCommand";
			SelectionName = "Begin If";
			CommandEnabled = true;
			CommandIcon = Resources.command_begin_if;
			ScopeStartCommand = true;

			//define parameter table
			v_ActionParameterTable = new DataTable
			{
				TableName = DateTime.Now.ToString("IfActionParamTable" + DateTime.Now.ToString("MMddyy.hhmmss"))
			};
			v_ActionParameterTable.Columns.Add("Parameter Name");
			v_ActionParameterTable.Columns.Add("Parameter Value");

			_recorderControl = new CommandItemControl();
			_recorderControl.Padding = new Padding(10, 0, 0, 0);
			_recorderControl.ForeColor = Color.AliceBlue;
			_recorderControl.Font = new Font("Segoe UI Semilight", 10);
			_recorderControl.Name = "guirecorder_helper";
			_recorderControl.CommandImage = Resources.command_camera;
			_recorderControl.CommandDisplay = "Element Recorder";
			_recorderControl.Hide();
		}

		public async override Tasks.Task RunCommand(object sender, ScriptAction parentCommand)
		{
			var engine = (IAutomationEngineInstance)sender;
			var ifResult = await CommandsHelper.DetermineStatementTruth(engine, v_IfActionType, v_ActionParameterTable);

			int startIndex, endIndex, elseIndex;
			if (parentCommand.AdditionalScriptCommands.Any(item => item.ScriptCommand is ElseCommand))
			{
				elseIndex = parentCommand.AdditionalScriptCommands.FindIndex(a => a.ScriptCommand is ElseCommand);

				if (ifResult)
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
			else if (ifResult)
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
				if ((engine.IsCancellationPending) || (engine.CurrentLoopCancelled) || (engine.CurrentLoopContinuing))
					return;

				await engine.ExecuteCommand(parentCommand.AdditionalScriptCommands[i]);
			}

		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			_actionDropdown = commandControls.CreateDropdownFor("v_IfActionType", this);
			RenderedControls.Add(commandControls.CreateDefaultLabelFor("v_IfActionType", this));
			RenderedControls.AddRange(commandControls.CreateUIHelpersFor("v_IfActionType", this, new Control[] { _actionDropdown }, editor));
			_actionDropdown.SelectionChangeCommitted += ifAction_SelectionChangeCommitted;
			RenderedControls.Add(_actionDropdown);

			_parameterControls = new List<Control>();
			_parameterControls.Add(commandControls.CreateDefaultLabelFor("v_ActionParameterTable", this));

			_recorderControl.Click += (sender, e) => ShowIfElementRecorder(sender, e, editor, commandControls);
			_parameterControls.Add(_recorderControl);

			_ifGridViewHelper = commandControls.CreateDefaultDataGridViewFor("v_ActionParameterTable", this);
			_ifGridViewHelper.AllowUserToAddRows = false;
			_ifGridViewHelper.AllowUserToDeleteRows = false;
			_ifGridViewHelper.MouseEnter += IfGridViewHelper_MouseEnter;

			_parameterControls.AddRange(commandControls.CreateUIHelpersFor("v_ActionParameterTable", this, new Control[] { _ifGridViewHelper }, editor));
			_parameterControls.Add(_ifGridViewHelper);

			RenderedControls.AddRange(_parameterControls);

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			switch (v_IfActionType)
			{
				case "Number Compare":
					string number1 = ((from rw in v_ActionParameterTable.AsEnumerable()
									  where rw.Field<string>("Parameter Name") == "Number1"
									  select rw.Field<string>("Parameter Value")).FirstOrDefault());
					string operand = ((from rw in v_ActionParameterTable.AsEnumerable()
									   where rw.Field<string>("Parameter Name") == "Operand"
									   select rw.Field<string>("Parameter Value")).FirstOrDefault());
					string number2 = ((from rw in v_ActionParameterTable.AsEnumerable()
									  where rw.Field<string>("Parameter Name") == "Number2"
									  select rw.Field<string>("Parameter Value")).FirstOrDefault());

					return $"If ('{number1}' {operand} '{number2}')";
				case "Date Compare":
					string date1 = ((from rw in v_ActionParameterTable.AsEnumerable()
									  where rw.Field<string>("Parameter Name") == "Date1"
									  select rw.Field<string>("Parameter Value")).FirstOrDefault());
					string operand2 = ((from rw in v_ActionParameterTable.AsEnumerable()
									   where rw.Field<string>("Parameter Name") == "Operand"
									   select rw.Field<string>("Parameter Value")).FirstOrDefault());
					string date2 = ((from rw in v_ActionParameterTable.AsEnumerable()
									  where rw.Field<string>("Parameter Name") == "Date2"
									  select rw.Field<string>("Parameter Value")).FirstOrDefault());

					return $"If ('{date1}' {operand2} '{date2}')";
				case "Text Compare":
					string text1 = ((from rw in v_ActionParameterTable.AsEnumerable()
									  where rw.Field<string>("Parameter Name") == "Text1"
									  select rw.Field<string>("Parameter Value")).FirstOrDefault());
					string operand3 = ((from rw in v_ActionParameterTable.AsEnumerable()
									   where rw.Field<string>("Parameter Name") == "Operand"
									   select rw.Field<string>("Parameter Value")).FirstOrDefault());
					string text2 = ((from rw in v_ActionParameterTable.AsEnumerable()
									  where rw.Field<string>("Parameter Name") == "Text2"
									  select rw.Field<string>("Parameter Value")).FirstOrDefault());

					return $"If ('{text1}' {operand3} '{text2}')";

				case "Has Value":
					string variableName = ((from rw in v_ActionParameterTable.AsEnumerable()
									  where rw.Field<string>("Parameter Name") == "Variable Name"
									  select rw.Field<string>("Parameter Value")).FirstOrDefault());

					return $"If (Variable '{variableName}' Has Value)";
				case "Is Numeric":
					string varName = ((from rw in v_ActionParameterTable.AsEnumerable()
											where rw.Field<string>("Parameter Name") == "Variable Name"
											select rw.Field<string>("Parameter Value")).FirstOrDefault());

					return $"If (Variable '{varName}' Is Numeric)";

				case "Error Occured":

					string lineNumber = ((from rw in v_ActionParameterTable.AsEnumerable()
										  where rw.Field<string>("Parameter Name") == "Line Number"
										  select rw.Field<string>("Parameter Value")).FirstOrDefault());

					return $"If (Error Occured on Line Number '{lineNumber}')";
				case "Error Did Not Occur":

					string lineNum = ((from rw in v_ActionParameterTable.AsEnumerable()
										  where rw.Field<string>("Parameter Name") == "Line Number"
										  select rw.Field<string>("Parameter Value")).FirstOrDefault());

					return $"If (Error Did Not Occur on Line Number '{lineNum}')";
				case "Window Name Exists":
				case "Active Window Name Is":

					string windowName = ((from rw in v_ActionParameterTable.AsEnumerable()
										  where rw.Field<string>("Parameter Name") == "Window Name"
										  select rw.Field<string>("Parameter Value")).FirstOrDefault());

					return $"If {v_IfActionType} [Window Name '{windowName}']";
				case "File Exists":

					string filePath = ((from rw in v_ActionParameterTable.AsEnumerable()
										where rw.Field<string>("Parameter Name") == "File Path"
										select rw.Field<string>("Parameter Value")).FirstOrDefault());

					string fileCompareType = ((from rw in v_ActionParameterTable.AsEnumerable()
											   where rw.Field<string>("Parameter Name") == "True When"
											   select rw.Field<string>("Parameter Value")).FirstOrDefault());

					if (fileCompareType == "It Does Not Exist")
						return $"If File Does Not Exist [File '{filePath}']";
					else
						return $"If File Exists [File '{filePath}']";

				case "Folder Exists":

					string folderPath = ((from rw in v_ActionParameterTable.AsEnumerable()
										  where rw.Field<string>("Parameter Name") == "Folder Path"
										  select rw.Field<string>("Parameter Value")).FirstOrDefault());

					string folderCompareType = ((from rw in v_ActionParameterTable.AsEnumerable()
											   where rw.Field<string>("Parameter Name") == "True When"
											   select rw.Field<string>("Parameter Value")).FirstOrDefault());

					if (folderCompareType == "It Does Not Exist")
						return $"If Folder Does Not [Folder '{folderPath}']";
					else
						return $"If Folder Exists [Folder '{folderPath}']";

				case "Web Element Exists":


					string parameterName = ((from rw in v_ActionParameterTable.AsEnumerable()
											 where rw.Field<string>("Parameter Name") == "Element Search Parameter"
											 select rw.Field<string>("Parameter Value")).FirstOrDefault());

					string searchMethod = ((from rw in v_ActionParameterTable.AsEnumerable()
											where rw.Field<string>("Parameter Name") == "Element Search Method"
											select rw.Field<string>("Parameter Value")).FirstOrDefault());

					string webElementCompareType = ((from rw in v_ActionParameterTable.AsEnumerable()
												 where rw.Field<string>("Parameter Name") == "True When"
												 select rw.Field<string>("Parameter Value")).FirstOrDefault());

					if (webElementCompareType == "It Does Not Exist")
						return $"If Web Element Does Not Exist [{searchMethod} '{parameterName}']";
					else
						return $"If Web Element Exists [{searchMethod} '{parameterName}']";

				case "GUI Element Exists":


					string guiWindowName = ((from rw in v_ActionParameterTable.AsEnumerable()
										 where rw.Field<string>("Parameter Name") == "Window Name"
										 select rw.Field<string>("Parameter Value")).FirstOrDefault());

					string guiSearch = ((from rw in v_ActionParameterTable.AsEnumerable()
											 where rw.Field<string>("Parameter Name") == "Element Search Parameter"
											 select rw.Field<string>("Parameter Value")).FirstOrDefault());

					string guiElementCompareType = ((from rw in v_ActionParameterTable.AsEnumerable()
													 where rw.Field<string>("Parameter Name") == "True When"
													 select rw.Field<string>("Parameter Value")).FirstOrDefault());

					if (guiElementCompareType == "It Does Not Exist")
						return $"If GUI Element Does Not Exist [Find '{guiSearch}' Element In '{guiWindowName}']";
					else
						return $"If GUI Element Exists [Find '{guiSearch}' Element In '{guiWindowName}']";

				case "Image Element Exists":
					string imageCompareType = (from rw in v_ActionParameterTable.AsEnumerable()
											   where rw.Field<string>("Parameter Name") == "True When"
											   select rw.Field<string>("Parameter Value")).FirstOrDefault();

					if (imageCompareType == "It Does Not Exist")
						return $"If Image Does Not Exist on Screen";
					else
						return $"If Image Exists on Screen";
				case "App Instance Exists":
					string instanceName = ((from rw in v_ActionParameterTable.AsEnumerable()
											 where rw.Field<string>("Parameter Name") == "Instance Name"
											 select rw.Field<string>("Parameter Value")).FirstOrDefault());

					string instanceCompareType = (from rw in v_ActionParameterTable.AsEnumerable()
											   where rw.Field<string>("Parameter Name") == "True When"
											   select rw.Field<string>("Parameter Value")).FirstOrDefault();

					if (instanceCompareType == "It Does Not Exist")
						return $"If App Instance Does Not Exist [Instance Name '{instanceName}']";
					else
						return $"If App Instance Exists [Instance Name '{instanceName}']";
				default:
					return "If ...";
			}
		}

		private void ifAction_SelectionChangeCommitted(object sender, EventArgs e)
		{
			DataGridView ifActionParameterBox = _ifGridViewHelper;

			BeginIfCommand cmd = this;
			DataTable actionParameters = cmd.v_ActionParameterTable;

			//sender is null when command is updating
			if (sender != null)
				actionParameters.Rows.Clear();

			DataGridViewComboBoxCell comparisonComboBox;

			//remove if exists            
			if (_recorderControl.Visible)
			{
				_recorderControl.Hide();
			}

			switch (_actionDropdown.SelectedItem)
			{
				case "Number Compare":

					ifActionParameterBox.Visible = true;

					if (sender != null)
					{
						actionParameters.Rows.Add("Number1", "");
						actionParameters.Rows.Add("Operand", "");
						actionParameters.Rows.Add("Number2", "");
						ifActionParameterBox.DataSource = actionParameters;
					}

					//combobox cell for Variable Name
					comparisonComboBox = new DataGridViewComboBoxCell();
					comparisonComboBox.Items.Add("is equal to");
					comparisonComboBox.Items.Add("is greater than");
					comparisonComboBox.Items.Add("is greater than or equal to");
					comparisonComboBox.Items.Add("is less than");
					comparisonComboBox.Items.Add("is less than or equal to");
					comparisonComboBox.Items.Add("is not equal to");

					//assign cell as a combobox
					ifActionParameterBox.Rows[1].Cells[1] = comparisonComboBox;

					break;
				case "Date Compare":

					ifActionParameterBox.Visible = true;

					if (sender != null)
					{
						actionParameters.Rows.Add("Date1", "");
						actionParameters.Rows.Add("Operand", "");
						actionParameters.Rows.Add("Date2", "");
						ifActionParameterBox.DataSource = actionParameters;
					}

					//combobox cell for Variable Name
					comparisonComboBox = new DataGridViewComboBoxCell();
					comparisonComboBox.Items.Add("is equal to");
					comparisonComboBox.Items.Add("is greater than");
					comparisonComboBox.Items.Add("is greater than or equal to");
					comparisonComboBox.Items.Add("is less than");
					comparisonComboBox.Items.Add("is less than or equal to");
					comparisonComboBox.Items.Add("is not equal to");

					//assign cell as a combobox
					ifActionParameterBox.Rows[1].Cells[1] = comparisonComboBox;

					break;
				case "Text Compare":

					ifActionParameterBox.Visible = true;

					if (sender != null)
					{
						actionParameters.Rows.Add("Text1", "");
						actionParameters.Rows.Add("Operand", "");
						actionParameters.Rows.Add("Text2", "");
						actionParameters.Rows.Add("Case Sensitive", "No");
						ifActionParameterBox.DataSource = actionParameters;
					}

					//combobox cell for Variable Name
					comparisonComboBox = new DataGridViewComboBoxCell();
					comparisonComboBox.Items.Add("contains");
					comparisonComboBox.Items.Add("does not contain");
					comparisonComboBox.Items.Add("is equal to");
					comparisonComboBox.Items.Add("is not equal to");

					//assign cell as a combobox
					ifActionParameterBox.Rows[1].Cells[1] = comparisonComboBox;

					comparisonComboBox = new DataGridViewComboBoxCell();
					comparisonComboBox.Items.Add("Yes");
					comparisonComboBox.Items.Add("No");

					//assign cell as a combobox
					ifActionParameterBox.Rows[3].Cells[1] = comparisonComboBox;

					break;
				case "Has Value":

					ifActionParameterBox.Visible = true;
					if (sender != null)
					{
						actionParameters.Rows.Add("Variable Name", "");
						ifActionParameterBox.DataSource = actionParameters;
					}

					break;
				case "Is Numeric":

					ifActionParameterBox.Visible = true;
					if (sender != null)
					{
						actionParameters.Rows.Add("Variable Name", "");
						ifActionParameterBox.DataSource = actionParameters;
					}

					break;
				case "Error Occured":

					ifActionParameterBox.Visible = true;
					if (sender != null)
					{
						actionParameters.Rows.Add("Line Number", "");
						ifActionParameterBox.DataSource = actionParameters;
					}

					break;
				case "Error Did Not Occur":

					ifActionParameterBox.Visible = true;

					if (sender != null)
					{
						actionParameters.Rows.Add("Line Number", "");
						ifActionParameterBox.DataSource = actionParameters;
					}

					break;
				case "Window Name Exists":
				case "Active Window Name Is":

					ifActionParameterBox.Visible = true;
					if (sender != null)
					{
						actionParameters.Rows.Add("Window Name", "");
						ifActionParameterBox.DataSource = actionParameters;
					}

					break;
				case "File Exists":

					ifActionParameterBox.Visible = true;
					if (sender != null)
					{
						actionParameters.Rows.Add("File Path", "");
						actionParameters.Rows.Add("True When", "It Does Exist");
						ifActionParameterBox.DataSource = actionParameters;
					}

					//combobox cell for Variable Name
					comparisonComboBox = new DataGridViewComboBoxCell();
					comparisonComboBox.Items.Add("It Does Exist");
					comparisonComboBox.Items.Add("It Does Not Exist");

					//assign cell as a combobox
					ifActionParameterBox.Rows[1].Cells[1] = comparisonComboBox;

					break;
				case "Folder Exists":

					ifActionParameterBox.Visible = true;

					if (sender != null)
					{
						actionParameters.Rows.Add("Folder Path", "");
						actionParameters.Rows.Add("True When", "It Does Exist");
						ifActionParameterBox.DataSource = actionParameters;
					}

					//combobox cell for Variable Name
					comparisonComboBox = new DataGridViewComboBoxCell();
					comparisonComboBox.Items.Add("It Does Exist");
					comparisonComboBox.Items.Add("It Does Not Exist");

					//assign cell as a combobox
					ifActionParameterBox.Rows[1].Cells[1] = comparisonComboBox;
					break;
				case "Web Element Exists":

					ifActionParameterBox.Visible = true;

					if (sender != null)
					{
						actionParameters.Rows.Add("Selenium Instance Name", "DefaultBrowser");
						actionParameters.Rows.Add("Element Search Method", "");
						actionParameters.Rows.Add("Element Search Parameter", "");
						actionParameters.Rows.Add("Timeout (Seconds)", "30");
						actionParameters.Rows.Add("True When", "It Does Exist");
						ifActionParameterBox.DataSource = actionParameters;
					}

					comparisonComboBox = new DataGridViewComboBoxCell();
					comparisonComboBox.Items.Add("It Does Exist");
					comparisonComboBox.Items.Add("It Does Not Exist");

					//assign cell as a combobox
					ifActionParameterBox.Rows[4].Cells[1] = comparisonComboBox;

					comparisonComboBox = new DataGridViewComboBoxCell();
					comparisonComboBox.Items.Add("XPath");
					comparisonComboBox.Items.Add("ID");
					comparisonComboBox.Items.Add("Name");
					comparisonComboBox.Items.Add("Tag Name");
					comparisonComboBox.Items.Add("Class Name");
					comparisonComboBox.Items.Add("CSS Selector");

					//assign cell as a combobox
					ifActionParameterBox.Rows[1].Cells[1] = comparisonComboBox;

					break;
				case "GUI Element Exists":

					ifActionParameterBox.Visible = true;
					if (sender != null)
					{
						actionParameters.Rows.Add("Window Name", "Current Window");
						actionParameters.Rows.Add("Element Search Method", "AutomationId");
						actionParameters.Rows.Add("Element Search Parameter", "");
						actionParameters.Rows.Add("Timeout (Seconds)", "30");
						actionParameters.Rows.Add("True When", "It Does Exist");
						ifActionParameterBox.DataSource = actionParameters;
					}

					comparisonComboBox = new DataGridViewComboBoxCell();
					comparisonComboBox.Items.Add("It Does Exist");
					comparisonComboBox.Items.Add("It Does Not Exist");

					//assign cell as a combobox
					ifActionParameterBox.Rows[4].Cells[1] = comparisonComboBox;

					var parameterName = new DataGridViewComboBoxCell();
					parameterName.Items.Add("AcceleratorKey");
					parameterName.Items.Add("AccessKey");
					parameterName.Items.Add("AutomationId");
					parameterName.Items.Add("ClassName");
					parameterName.Items.Add("FrameworkId");
					parameterName.Items.Add("HasKeyboardFocus");
					parameterName.Items.Add("HelpText");
					parameterName.Items.Add("IsContentElement");
					parameterName.Items.Add("IsControlElement");
					parameterName.Items.Add("IsEnabled");
					parameterName.Items.Add("IsKeyboardFocusable");
					parameterName.Items.Add("IsOffscreen");
					parameterName.Items.Add("IsPassword");
					parameterName.Items.Add("IsRequiredForForm");
					parameterName.Items.Add("ItemStatus");
					parameterName.Items.Add("ItemType");
					parameterName.Items.Add("LocalizedControlType");
					parameterName.Items.Add("Name");
					parameterName.Items.Add("NativeWindowHandle");
					parameterName.Items.Add("ProcessID");

					//assign cell as a combobox
					ifActionParameterBox.Rows[1].Cells[1] = parameterName;

					_recorderControl.Show();

					break;
				case "Image Element Exists":
					ifActionParameterBox.Visible = true;

					if (sender != null)
					{
						actionParameters.Rows.Add("Captured Image Variable", "");
						actionParameters.Rows.Add("Accuracy (0-1)", "0.8");
						actionParameters.Rows.Add("True When", "It Does Exist");
						actionParameters.Rows.Add("Timeout (Seconds)", "30");
						ifActionParameterBox.DataSource = actionParameters;
					}

					comparisonComboBox = new DataGridViewComboBoxCell();
					comparisonComboBox.Items.Add("It Does Exist");
					comparisonComboBox.Items.Add("It Does Not Exist");

					//assign cell as a combobox
					ifActionParameterBox.Rows[2].Cells[1] = comparisonComboBox;
					break;
				case "App Instance Exists":
					ifActionParameterBox.Visible = true;

					if (sender != null)
					{
						actionParameters.Rows.Add("Instance Name", "");
						actionParameters.Rows.Add("True When", "It Does Exist");
						ifActionParameterBox.DataSource = actionParameters;
					}

					comparisonComboBox = new DataGridViewComboBoxCell();
					comparisonComboBox.Items.Add("It Does Exist");
					comparisonComboBox.Items.Add("It Does Not Exist");

					//assign cell as a combobox
					ifActionParameterBox.Rows[1].Cells[1] = comparisonComboBox;
					break;
				default:
					break;
			}

			ifActionParameterBox.Columns[0].ReadOnly = true;
		}

		private void IfGridViewHelper_MouseEnter(object sender, EventArgs e)
		{
			try
			{
				ifAction_SelectionChangeCommitted(null, null);
			}
			catch (Exception)
			{
				ifAction_SelectionChangeCommitted(sender, e);
			}           
		}

		private void ShowIfElementRecorder(object sender, EventArgs e, IfrmCommandEditor editor, ICommandControls commandControls)
		{
			var result = commandControls.ShowConditionElementRecorder(sender, e, editor);

			_ifGridViewHelper.Rows[0].Cells[1].Value = result.Item1;
			_ifGridViewHelper.Rows[2].Cells[1].Value = result.Item2;
		}
	}
}
