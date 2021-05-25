﻿using Newtonsoft.Json;
using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Properties;
using OpenBots.Core.UI.Controls;
using OpenBots.Core.User32;
using OpenBots.Core.Utilities.CommandUtilities;
using OpenBots.Core.Utilities.CommonUtilities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Security;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;

namespace OpenBots.Commands.Input
{
    [Serializable]
	[Category("Input Commands")]
	[Description("This Command automates an element in a targeted window.")]
	public class UIAutomationCommand : ScriptCommand, IUIAutomationCommand
	{
		[Required]
		[DisplayName("Window Name")]
		[Description("Select the name of the window to automate.")]
		[SampleUsage("\"Untitled - Notepad\" || \"Current Window\" || vWindow")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[Editor("CaptureWindowHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_WindowName { get; set; }

		[Required]
		[DisplayName("Element Action")]
		[PropertyUISelectionOption("Click Element")]
		[PropertyUISelectionOption("Set Text")]
		[PropertyUISelectionOption("Set Secure Text")]
		[PropertyUISelectionOption("Get Text")]
		[PropertyUISelectionOption("Clear Element")]
		[PropertyUISelectionOption("Get Value From Element")]
		[PropertyUISelectionOption("Wait For Element To Exist")]
		[PropertyUISelectionOption("Element Exists")]
		[Description("Select the appropriate corresponding action to take once the element has been located.")]
		[SampleUsage("")]
		[Remarks("Selecting this field changes the parameters required in the following step.")]
		public string v_AutomationType { get; set; }

		[Required]
		[DisplayName("Element Search Parameter")]
		[Description("Use the Element Recorder to generate a listing of potential search parameters.")]
		[SampleUsage("[ AutomationId | \"Name\" ]")]
		[Remarks("Once you have clicked on a valid window the search parameters will be populated. Select a single parameter to find the element.")]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public DataTable v_UIASearchParameters { get; set; }

		[Required]
		[DisplayName("Action Parameters")]
		[Description("Action Parameters will be determined based on the action settings selected.")]
		[SampleUsage("\"data\" || vData || vOutputVariable")]
		[Remarks("Action Parameters range from adding offset coordinates to specifying a variable to apply element text to.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(SecureString), typeof(string), typeof(bool), typeof(int) })]
		public DataTable v_UIAActionParameters { get; set; }

		[Required]
		[DisplayName("Timeout (Seconds)")]
		[Description("Specify how many seconds to wait before throwing an exception.")]
		[SampleUsage("30 || vSeconds")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(int) })]
		public string v_Timeout { get; set; }

		[JsonIgnore]
		[Browsable(false)]
		private ComboBox _automationTypeControl;

		[JsonIgnore]
		[Browsable(false)]
		private DataGridView _searchParametersGridViewHelper;

		[JsonIgnore]
		[Browsable(false)]
		private DataGridView _actionParametersGridViewHelper;

		[JsonIgnore]
		[Browsable(false)]
		private List<Control> _actionParametersControls;

		public UIAutomationCommand()
		{
			CommandName = "UIAutomationCommand";
			SelectionName = "UI Automation";
			CommandEnabled = true;
			CommandIcon = Resources.command_input;

			//set up search parameter table
			v_UIASearchParameters = new DataTable();
			v_UIASearchParameters.Columns.Add("Enabled");
			v_UIASearchParameters.Columns.Add("Parameter Name");
			v_UIASearchParameters.Columns.Add("Parameter Value");
			v_UIASearchParameters.TableName = DateTime.Now.ToString("UIASearchParamTable" + DateTime.Now.ToString("MMddyy.hhmmss"));

			v_UIAActionParameters = new DataTable();
			v_UIAActionParameters.Columns.Add("Parameter Name");
			v_UIAActionParameters.Columns.Add("Parameter Value");
			v_UIAActionParameters.TableName = DateTime.Now.ToString("UIAActionParamTable" + DateTime.Now.ToString("MMddyy.hhmmss"));

			v_WindowName = "\"Current Window\"";
			v_Timeout = "30";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var vTimeout = (int)await v_Timeout.EvaluateCode(engine);

			//create variable window name
			var variableWindowName = (string)await v_WindowName.EvaluateCode(engine);
			if (variableWindowName == "Current Window")
				variableWindowName = User32Functions.GetActiveWindowTitle();			

			AutomationElement requiredHandle = null;

			var timeToEnd = DateTime.Now.AddSeconds(vTimeout);
			while (timeToEnd >= DateTime.Now)
			{
				try
				{
					if (engine.IsCancellationPending)
						break;

					requiredHandle = await CommandsHelper.SearchForGUIElement(engine, v_UIASearchParameters, variableWindowName);

					if (requiredHandle == null)
						throw new Exception("Element Not Yet Found... ");

					break;
				}
				catch (Exception)
				{
					engine.ReportProgress("Element Not Yet Found... " + (timeToEnd - DateTime.Now).Seconds + "s remain");
					Thread.Sleep(500);
				}
			}

			switch (v_AutomationType)
			{
				//determine element click type
				case "Click Element":
					//if handle was not found
					if (requiredHandle == null)
						throw new Exception("Element was not found in window '" + variableWindowName + "'");
					//create search params
					var clickType = (from rw in v_UIAActionParameters.AsEnumerable()
									 where rw.Field<string>("Parameter Name") == "Click Type"
									 select rw.Field<string>("Parameter Value")).FirstOrDefault();

					//get x adjust
					var xAdjust = (from rw in v_UIAActionParameters.AsEnumerable()
								   where rw.Field<string>("Parameter Name") == "X Adjustment"
								   select rw.Field<string>("Parameter Value")).FirstOrDefault();

					//get y adjust
					var yAdjust = (from rw in v_UIAActionParameters.AsEnumerable()
								   where rw.Field<string>("Parameter Name") == "Y Adjustment"
								   select rw.Field<string>("Parameter Value")).FirstOrDefault();

					int xAdjustInt;
					int yAdjustInt;

					//parse to int
					if (!string.IsNullOrEmpty(xAdjust))
						xAdjustInt = (int)await xAdjust.EvaluateCode(engine);
					else
						xAdjustInt = 0;

					if (!string.IsNullOrEmpty(yAdjust))
						yAdjustInt = (int)await yAdjust.EvaluateCode(engine);
					else
						yAdjustInt = 0;

					//get clickable point
					var newPoint = requiredHandle.GetClickablePoint();

					//send mousemove command
					User32Functions.SendMouseMove(Convert.ToInt32(newPoint.X) + xAdjustInt, Convert.ToInt32(newPoint.Y) + yAdjustInt, clickType);

					break;
				case "Set Text":
					string textToSet = (from rw in v_UIAActionParameters.AsEnumerable()
										where rw.Field<string>("Parameter Name") == "Text To Set"
										select rw.Field<string>("Parameter Value")).FirstOrDefault();


					string clearElement = (from rw in v_UIAActionParameters.AsEnumerable()
										   where rw.Field<string>("Parameter Name") == "Clear Element Before Setting Text"
										   select rw.Field<string>("Parameter Value")).FirstOrDefault();

					if (clearElement == null)
						clearElement = "No";

					textToSet = (string)await textToSet.EvaluateCode(engine);

					if (requiredHandle.Current.IsEnabled && requiredHandle.Current.IsKeyboardFocusable)
					{
						object valuePattern = null;
						if (!requiredHandle.TryGetCurrentPattern(ValuePattern.Pattern, out valuePattern))
						{
							//The control does not support ValuePattern Using keyboard input
							// Set focus for input functionality and begin.
							requiredHandle.SetFocus();

							// Pause before sending keyboard input.
							Thread.Sleep(100);

							if (clearElement.ToLower() == "yes")
							{
								// Delete existing content in the control and insert new content.
								SendKeys.SendWait("^{HOME}");   // Move to start of control
								SendKeys.SendWait("^+{END}");   // Select everything
								SendKeys.SendWait("{DEL}");     // Delete selection
							}
							SendKeys.SendWait(textToSet);
						}
						else
						{
							if (clearElement.ToLower() == "no")
							{
								string currentText;
								object tPattern = null;
								if (requiredHandle.TryGetCurrentPattern(TextPattern.Pattern, out tPattern))
								{
									var textPattern = (TextPattern)tPattern;
									// often there is an extra '\r' hanging off the end.
									currentText = textPattern.DocumentRange.GetText(-1).TrimEnd('\r').ToString(); 
								}
								else
									currentText = requiredHandle.Current.Name.ToString();

								textToSet = currentText + textToSet;
							}
							requiredHandle.SetFocus();
							((ValuePattern)valuePattern).SetValue(textToSet);
						}
					}
					break;
				case "Set Secure Text":
					string secureString = (from rw in v_UIAActionParameters.AsEnumerable()
										   where rw.Field<string>("Parameter Name") == "Secure String Variable"
										   select rw.Field<string>("Parameter Value")).FirstOrDefault();

					string _clearElement = (from rw in v_UIAActionParameters.AsEnumerable()
											where rw.Field<string>("Parameter Name") == "Clear Element Before Setting Text"
											select rw.Field<string>("Parameter Value")).FirstOrDefault();

					var secureStrVariable = (SecureString)await secureString.EvaluateCode(engine);

					secureString = secureStrVariable.ConvertSecureStringToString();

					if (_clearElement == null)
						_clearElement = "No";

					if (requiredHandle.Current.IsEnabled && requiredHandle.Current.IsKeyboardFocusable)
					{
						object valuePattern = null;
						if (!requiredHandle.TryGetCurrentPattern(ValuePattern.Pattern, out valuePattern))
						{
							//The control does not support ValuePattern Using keyboard input
							// Set focus for input functionality and begin.
							requiredHandle.SetFocus();

							// Pause before sending keyboard input.
							Thread.Sleep(100);

							if (_clearElement.ToLower() == "yes")
							{
								// Delete existing content in the control and insert new content.
								SendKeys.SendWait("^{HOME}");   // Move to start of control
								SendKeys.SendWait("^+{END}");   // Select everything
								SendKeys.SendWait("{DEL}");     // Delete selection
							}
							SendKeys.SendWait(secureString);
						}
						else
						{
							if (_clearElement.ToLower() == "no")
							{
								string currentText;
								object tPattern = null;
								if (requiredHandle.TryGetCurrentPattern(TextPattern.Pattern, out tPattern))
								{
									var textPattern = (TextPattern)tPattern;
									currentText = textPattern.DocumentRange.GetText(-1).TrimEnd('\r').ToString(); // often there is an extra '\r' hanging off the end.
								}
								else
									currentText = requiredHandle.Current.Name.ToString();

								secureString = currentText + secureString;
							}
							requiredHandle.SetFocus();
							((ValuePattern)valuePattern).SetValue(secureString);
						}
					}
					break;
				case "Clear Element":
					if (requiredHandle.Current.IsEnabled && requiredHandle.Current.IsKeyboardFocusable)
					{
						object valuePattern = null;
						if (!requiredHandle.TryGetCurrentPattern(ValuePattern.Pattern, out valuePattern))
						{
							//The control does not support ValuePattern Using keyboard input
							// Set focus for input functionality and begin.
							requiredHandle.SetFocus();

							// Pause before sending keyboard input.
							Thread.Sleep(100);

							// Delete existing content in the control and insert new content.
							SendKeys.SendWait("^{HOME}");   // Move to start of control
							SendKeys.SendWait("^+{END}");   // Select everything
							SendKeys.SendWait("{DEL}");     // Delete selection
						}
						else
						{
							requiredHandle.SetFocus();
							((ValuePattern)valuePattern).SetValue("");
						}
					}
					break;
				case "Get Text":
				//if element exists type
				case "Element Exists":
					//Variable Name
					var applyToVariable = (from rw in v_UIAActionParameters.AsEnumerable()
										   where rw.Field<string>("Parameter Name") == "Variable Name"
										   select rw.Field<string>("Parameter Value")).FirstOrDefault();

					//declare search result
					dynamic searchResult;
					if (v_AutomationType == "Get Text")
					{
						//string currentText;
						object tPattern = null;
						if (requiredHandle.TryGetCurrentPattern(TextPattern.Pattern, out tPattern))
						{
							var textPattern = (TextPattern)tPattern;
							searchResult = textPattern.DocumentRange.GetText(-1).TrimEnd('\r').ToString(); // often there is an extra '\r' hanging off the end.
						}
						else
							searchResult = requiredHandle.Current.Name.ToString();

						((string)searchResult).SetVariableValue(engine, applyToVariable);
					}

					else if (v_AutomationType == "Element Exists")
					{
						//determine search result
						if (requiredHandle == null)
							searchResult = false;
						else
							searchResult = true;

						((bool)searchResult).SetVariableValue(engine, applyToVariable);
					}
					
					break;
				case "Wait For Element To Exist":
					if (requiredHandle == null)
					{
						throw new Exception($"Element was not found in the allotted time!");
					}
					break;

				case "Get Value From Element":
					if (requiredHandle == null)
						throw new Exception("Element was not found in window '" + variableWindowName + "'");
					//get value from property
					var propertyName = (from rw in v_UIAActionParameters.AsEnumerable()
										where rw.Field<string>("Parameter Name") == "Get Value From"
										select rw.Field<string>("Parameter Value")).FirstOrDefault();

					//Variable Name
					var applyToVariable2 = (from rw in v_UIAActionParameters.AsEnumerable()
										   where rw.Field<string>("Parameter Name") == "Variable Name"
										   select rw.Field<string>("Parameter Value")).FirstOrDefault();

					//get required value
					var requiredValue = requiredHandle.Current.GetType().GetRuntimeProperty(propertyName)?.GetValue(requiredHandle.Current).ToString();

					//store into variable
					((object)requiredValue).SetVariableValue(engine, applyToVariable2);
					break;
				default:
					throw new NotImplementedException("Automation type '" + v_AutomationType + "' not supported.");
			}
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			//create search param grid
			_searchParametersGridViewHelper = commandControls.CreateDefaultDataGridViewFor("v_UIASearchParameters", this);

			DataGridViewCheckBoxColumn enabled = new DataGridViewCheckBoxColumn();
			enabled.HeaderText = "Enabled";
			enabled.DataPropertyName = "Enabled";
			_searchParametersGridViewHelper.Columns.Add(enabled);

			DataGridViewTextBoxColumn propertyName = new DataGridViewTextBoxColumn();
			propertyName.HeaderText = "Parameter Name";
			propertyName.DataPropertyName = "Parameter Name";
			propertyName.ReadOnly = true;
			_searchParametersGridViewHelper.Columns.Add(propertyName);

			DataGridViewTextBoxColumn propertyValue = new DataGridViewTextBoxColumn();
			propertyValue.HeaderText = "Parameter Value";
			propertyValue.DataPropertyName = "Parameter Value";

			_searchParametersGridViewHelper.Columns.Add(propertyValue);
			_searchParametersGridViewHelper.AllowUserToAddRows = false;
			_searchParametersGridViewHelper.AllowUserToDeleteRows = false;
			_searchParametersGridViewHelper.MouseEnter += ActionParametersGridViewHelper_MouseEnter;

			//create actions
			_actionParametersGridViewHelper = commandControls.CreateDefaultDataGridViewFor("v_UIAActionParameters", this);
			_actionParametersGridViewHelper.AllowUserToAddRows = false;
			_actionParametersGridViewHelper.AllowUserToDeleteRows = false;
			_actionParametersGridViewHelper.MouseEnter += ActionParametersGridViewHelper_MouseEnter;

			propertyName = new DataGridViewTextBoxColumn();
			propertyName.HeaderText = "Parameter Name";
			propertyName.DataPropertyName = "Parameter Name";
			_actionParametersGridViewHelper.Columns.Add(propertyName);

			propertyValue = new DataGridViewTextBoxColumn();
			propertyValue.HeaderText = "Parameter Value";
			propertyValue.DataPropertyName = "Parameter Value";
			_actionParametersGridViewHelper.Columns.Add(propertyValue);

			//create helper control
			CommandItemControl helperControl = new CommandItemControl("UIRecorder", Resources.command_camera, "UI Element Recorder");
			helperControl.Click += (sender, e) => ShowRecorder(sender, e, commandControls);

			//window name
			RenderedControls.AddRange(commandControls.CreateDefaultWindowControlGroupFor("v_WindowName", this, editor));

			//automation type
			var automationTypeGroup = commandControls.CreateDefaultDropdownGroupFor("v_AutomationType", this, editor);
			_automationTypeControl = (ComboBox)automationTypeGroup.Where(f => f is ComboBox).FirstOrDefault();
			_automationTypeControl.SelectionChangeCommitted += UIAType_SelectionChangeCommitted;
			RenderedControls.AddRange(automationTypeGroup);

			//create search parameters   
			RenderedControls.Add(commandControls.CreateDefaultLabelFor("v_UIASearchParameters", this));
			RenderedControls.Add(helperControl);
			RenderedControls.Add(_searchParametersGridViewHelper);

			//create action parameters
			_actionParametersControls = new List<Control>();
			_actionParametersControls.Add(commandControls.CreateDefaultLabelFor("v_UIAActionParameters", this));
			_actionParametersControls.AddRange(commandControls.CreateUIHelpersFor("v_UIAActionParameters", this, new Control[] { _actionParametersGridViewHelper }, editor));
			_actionParametersControls.Add(_actionParametersGridViewHelper);
			RenderedControls.AddRange(_actionParametersControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_Timeout", this, editor));

			return RenderedControls;
		}
		
		public override string GetDisplayValue()
		{
			var applyToVariable = (from rw in v_UIAActionParameters.AsEnumerable()
								   where rw.Field<string>("Parameter Name") == "Variable Name"
								   select rw.Field<string>("Parameter Value")).FirstOrDefault();

			switch (v_AutomationType)
			{
				case "Click Element":
					//create search params
					var clickType = (from rw in v_UIAActionParameters.AsEnumerable()
									 where rw.Field<string>("Parameter Name") == "Click Type"
									 select rw.Field<string>("Parameter Value")).FirstOrDefault();

					return base.GetDisplayValue() + $" [{clickType} Element in Window '{v_WindowName}']";

				case "Set Text":
				case "Set Secure Text":
					var textToSet = (from rw in v_UIAActionParameters.AsEnumerable()
									 where rw.Field<string>("Parameter Name") == "Text To Set"
									 select rw.Field<string>("Parameter Value")).FirstOrDefault();
					return base.GetDisplayValue() + $" [{v_AutomationType} '{textToSet}' in Element in Window '{v_WindowName}']";
				  
				case "Get Text":
				case "Element Exists":          
					return base.GetDisplayValue() + $" ['{v_AutomationType}' in Window '{v_WindowName}' - Store Result in '{applyToVariable}']";
			   
				case "Get Value From Element":          
					//get value from property
					var propertyName = (from rw in v_UIAActionParameters.AsEnumerable()
										where rw.Field<string>("Parameter Name") == "Get Value From"
										select rw.Field<string>("Parameter Value")).FirstOrDefault();

					return base.GetDisplayValue() + $" [Get Value From Element '{propertyName}' in Window '{v_WindowName}' - Store Value in '{applyToVariable}']";

				default:
					return base.GetDisplayValue() + $" [{v_AutomationType} in Window '{v_WindowName}']";
			}
		}

		private void ActionParametersGridViewHelper_MouseEnter(object sender, EventArgs e)
		{
			UIAType_SelectionChangeCommitted(null, null);
		}

		public void ShowRecorder(object sender, EventArgs e, ICommandControls commandControls)
		{
			//get command reference
			//create recorder
			IfrmAdvancedUIElementRecorder newElementRecorder = commandControls.CreateAdvancedUIElementRecorderForm();
			newElementRecorder.WindowName = RenderedControls[3].Text;
			newElementRecorder.SearchParameters = v_UIASearchParameters;
			newElementRecorder.chkStopOnClick.Checked = true;

			//show form
			((Form)newElementRecorder).ShowDialog();

			//window name combobox
			RenderedControls[3].Text = $"\"{newElementRecorder.WindowName}\"";

			v_UIASearchParameters.Rows.Clear();

			foreach (DataRow rw in newElementRecorder.SearchParameters.Rows)
				v_UIASearchParameters.ImportRow(rw);
				
			_searchParametersGridViewHelper.DataSource = v_UIASearchParameters;
			_searchParametersGridViewHelper.Refresh();
		}

		public void UIAType_SelectionChangeCommitted(object sender, EventArgs e)
		{
			UIAutomationCommand cmd = this;
			DataTable actionParameters = cmd.v_UIAActionParameters;

			if (sender != null)
				actionParameters.Rows.Clear();

			switch (_automationTypeControl.SelectedItem)
			{
				case "Click Element":
					foreach (var ctrl in _actionParametersControls)
						ctrl.Show();

					var mouseClickBox = new DataGridViewComboBoxCell();
					mouseClickBox.Items.Add("Left Click");
					mouseClickBox.Items.Add("Middle Click");
					mouseClickBox.Items.Add("Right Click");
					mouseClickBox.Items.Add("Left Down");
					mouseClickBox.Items.Add("Middle Down");
					mouseClickBox.Items.Add("Right Down");
					mouseClickBox.Items.Add("Left Up");
					mouseClickBox.Items.Add("Middle Up");
					mouseClickBox.Items.Add("Right Up");
					mouseClickBox.Items.Add("Double Left Click");

					if (sender != null)
					{
						actionParameters.Rows.Add("Click Type", "");
						actionParameters.Rows.Add("X Adjustment", 0);
						actionParameters.Rows.Add("Y Adjustment", 0);
					}

					if (_actionParametersGridViewHelper.Rows.Count > 0)
						_actionParametersGridViewHelper.Rows[0].Cells[1] = mouseClickBox;
					break;
				case "Set Text":
					foreach (var ctrl in _actionParametersControls)
						ctrl.Show();

					if (sender != null)
					{
						actionParameters.Rows.Add("Text To Set");
						actionParameters.Rows.Add("Clear Element Before Setting Text");
					}

					DataGridViewComboBoxCell comparisonComboBox = new DataGridViewComboBoxCell();
					comparisonComboBox.Items.Add("Yes");
					comparisonComboBox.Items.Add("No");

					//assign cell as a combobox
					if (sender != null)
						_actionParametersGridViewHelper.Rows[1].Cells[1].Value = "No";

					if (_actionParametersGridViewHelper.Rows.Count > 1)
						_actionParametersGridViewHelper.Rows[1].Cells[1] = comparisonComboBox;

					break;

				case "Set Secure Text":
					foreach (var ctrl in _actionParametersControls)
						ctrl.Show();

					if (sender != null)
					{
						actionParameters.Rows.Add("Secure String Variable");
						actionParameters.Rows.Add("Clear Element Before Setting Text");
					}

					DataGridViewComboBoxCell _comparisonComboBox = new DataGridViewComboBoxCell();
					_comparisonComboBox.Items.Add("Yes");
					_comparisonComboBox.Items.Add("No");

					//assign cell as a combobox
					if (sender != null)
						_actionParametersGridViewHelper.Rows[1].Cells[1].Value = "No";

					if (_actionParametersGridViewHelper.Rows.Count > 1)
						_actionParametersGridViewHelper.Rows[1].Cells[1] = _comparisonComboBox;
					break;
				case "Get Text":
				case "Element Exists":
					foreach (var ctrl in _actionParametersControls)
						ctrl.Show();

					if (sender != null)
						actionParameters.Rows.Add("Variable Name", "");
					break;
				case "Clear Element":
					foreach (var ctrl in _actionParametersControls)
						ctrl.Hide();

					break;
				case "Wait For Element To Exist":
					foreach (var ctrl in _actionParametersControls)
						ctrl.Hide();

					break;
				case "Get Value From Element":
					foreach (var ctrl in _actionParametersControls)
						ctrl.Show();

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

					if (sender != null)
					{
						actionParameters.Rows.Add("Get Value From", "");
						actionParameters.Rows.Add("Variable Name", "");
						_actionParametersGridViewHelper.Refresh();
						try
						{
							_actionParametersGridViewHelper.Rows[0].Cells[1] = parameterName;
						}
						catch (Exception ex)
						{
							MessageBox.Show("Unable to select first row, second cell to apply '" + parameterName + "': " + ex.ToString());
						}
					}
					break;
				default:
					break;
			}
			_actionParametersGridViewHelper.Columns[0].ReadOnly = true;
			_actionParametersGridViewHelper.DataSource = v_UIAActionParameters;
		}
	}
}