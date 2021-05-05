﻿using HtmlAgilityPack;
using Newtonsoft.Json;
using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Properties;
using OpenBots.Core.UI.Controls;
using OpenBots.Core.User32;
using OpenBots.Core.Utilities.CommandUtilities;
using OpenBots.Core.Utilities.CommonUtilities;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Security;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace OpenBots.Commands.WebBrowser
{
	[Serializable]
	[Category("Web Browser Commands")]
	[Description("This command performs an element action in a Selenium web browser session.")]
	public class SeleniumElementActionCommand : ScriptCommand, ISeleniumElementActionCommand
	{
		[Required]
		[DisplayName("Browser Instance Name")]
		[Description("Enter the unique instance that was specified in the **Create Browser** command.")]
		[SampleUsage("MyBrowserInstance")]
		[Remarks("Failure to enter the correct instance name or failure to first call the **Create Browser** command will cause an error.")]
		[CompatibleTypes(new Type[] { typeof(IWebDriver) })]
		public string v_InstanceName { get; set; }

		[Required]
		[DisplayName("Element Search Parameter")]
		[Description("Use the Element Recorder to generate a listing of potential search parameters." + 
			"Select the specific search type(s) that you want to use to isolate the element on the web page.")]
		[SampleUsage("XPath : \"//*[@id=\'features\']/div[2]/div/h2/div[\" + var1 + \"]/div\"" +
				 "\n\tID: \"1\"" +
				 "\n\tName: \"my\" + var2 + \"Name\"" +
				 "\n\tTag Name: \"h1\"" +
				 "\n\tClass Name: \"myClass\"" +
				 "\n\tCSS Selector: \"[attribute=value]\"" +
				 "\n\tLink Text: \"https://www.mylink.com/\"")]
		[Remarks("If multiple parameters are enabled, an attempt will be made to find the element(s) that match(es) all the selected parameters.")]
		[Editor("ShowElementHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public DataTable v_SeleniumSearchParameters { get; set; }

		[Required]
		[DisplayName("Element Search Option")]
		[PropertyUISelectionOption("Find Element")]
		[PropertyUISelectionOption("Find Elements")]
		[Description("Indicate whether to search for a single or multiple elements.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_SeleniumSearchOption { get; set; }

		[Required]
		[DisplayName("Element Action")]
		[PropertyUISelectionOption("Invoke Click")]
		[PropertyUISelectionOption("Left Click")]
		[PropertyUISelectionOption("Right Click")]
		[PropertyUISelectionOption("Middle Click")]
		[PropertyUISelectionOption("Double Left Click")]
		[PropertyUISelectionOption("Set Text")]
		[PropertyUISelectionOption("Set Secure Text")]
		[PropertyUISelectionOption("Get Text")]
		[PropertyUISelectionOption("Get Table")]
		[PropertyUISelectionOption("Get Count")]
		[PropertyUISelectionOption("Get Options")]
		[PropertyUISelectionOption("Get Attribute")]
		[PropertyUISelectionOption("Get Matching Element(s)")]
		[PropertyUISelectionOption("Clear Element")]
		[PropertyUISelectionOption("Wait For Element To Exist")]
		[PropertyUISelectionOption("Element Exists")]
		[PropertyUISelectionOption("Switch to frame")]
		[PropertyUISelectionOption("Select Option")]
		[Description("Select the appropriate corresponding action to take once the element has been located.")]
		[SampleUsage("")]
		[Remarks("Selecting this field changes the parameters required in the following step.")]
		public string v_SeleniumElementAction { get; set; }

		[Required]
		[DisplayName("Action Parameters")]
		[Description("Action Parameters will be determined based on the action settings selected.")]
		[SampleUsage("\"data\" || vData || vOutputVariable")]
		[Remarks("Action Parameters range from adding offset coordinates to specifying a variable to apply element text to.\n"+
				 "Advanced keystrokes may be set the following way: Hello[tab]World[enter]")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(SecureString), typeof(IWebElement), typeof(List<IWebElement>), typeof(List<string>), 
			typeof(DataTable), typeof(string), typeof(bool) })]
		public DataTable v_WebActionParameterTable { get; set; }

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
		private DataGridView _actionParametersGridViewHelper;

		[JsonIgnore]
		[Browsable(false)]
		private ComboBox _elementActionDropdown;

		[JsonIgnore]
		[Browsable(false)]
		private List<Control> _actionParametersControls;

		[JsonIgnore]
		[Browsable(false)]
		private DataGridView _searchParametersGridViewHelper;

		public SeleniumElementActionCommand()
		{
			CommandName = "SeleniumElementActionCommand";
			SelectionName = "Element Action";
			v_InstanceName = "DefaultBrowser";
			CommandEnabled = true;
			CommandIcon = Resources.command_web;

			v_SeleniumSearchOption = "Find Element";
			v_Timeout = "30";

			v_WebActionParameterTable = new DataTable
			{
				TableName = "WebActionParamTable" + DateTime.Now.ToString("MMddyy.hhmmss")
			};
			v_WebActionParameterTable.Columns.Add("Parameter Name");
			v_WebActionParameterTable.Columns.Add("Parameter Value");

			//set up search parameter table
			v_SeleniumSearchParameters = new DataTable();
			v_SeleniumSearchParameters.Columns.Add("Enabled");
			v_SeleniumSearchParameters.Columns.Add("Parameter Name");
			v_SeleniumSearchParameters.Columns.Add("Parameter Value");
			v_SeleniumSearchParameters.TableName = DateTime.Now.ToString("v_SeleniumSearchParameters" + DateTime.Now.ToString("MMddyy.hhmmss"));		
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;

			var vTimeout = (int)await v_Timeout.EvaluateCode(engine);
			var seleniumSearchParamRows = (from rw in v_SeleniumSearchParameters.AsEnumerable()
									   where rw.Field<string>("Enabled") == "True" &&
									   rw.Field<string>("Parameter Value").ToString() != ""
									   select rw.ItemArray.Cast<string>().ToArray()).ToList();

			var browserObject = v_InstanceName.GetAppInstance(engine);
			var seleniumInstance = (IWebDriver)browserObject;
			dynamic element = await CommandsHelper.FindElement(engine, seleniumInstance, seleniumSearchParamRows, v_SeleniumSearchOption, vTimeout);

			if (element == null && v_SeleniumElementAction != "Element Exists")
				throw new ElementNotVisibleException("Unable to find element within the provided time limit");

			if (v_SeleniumElementAction.Contains("Click"))
			{
				int seleniumWindowHeightY = seleniumInstance.Manage().Window.Size.Height;
				int elementPositionY = ((IWebElement)element).Location.Y;
				if (elementPositionY > seleniumWindowHeightY)
				{
					string scroll = string.Format("window.scroll(0, {0})", elementPositionY);
					IJavaScriptExecutor js = browserObject as IJavaScriptExecutor;
					js.ExecuteScript(scroll);
				}                
			}
			Actions actions = new Actions(seleniumInstance);

			switch (v_SeleniumElementAction)
			{				
				case "Invoke Click":
					((IWebElement)element).Click();                   
					break;

				case "Left Click":
					actions.Click((IWebElement)element).Perform();
					break;
				case "Right Click":                   
					actions.ContextClick((IWebElement)element).Perform();
					break;
				case "Middle Click":
					string userXAdjustString = (from rw in v_WebActionParameterTable.AsEnumerable()
													   where rw.Field<string>("Parameter Name") == "X Adjustment"
													   select rw.Field<string>("Parameter Value")).FirstOrDefault();
					int userXAdjust = (int)await userXAdjustString.EvaluateCode(engine);

					string userYAdjustString = (from rw in v_WebActionParameterTable.AsEnumerable()
													   where rw.Field<string>("Parameter Name") == "Y Adjustment"
													   select rw.Field<string>("Parameter Value")).FirstOrDefault();
					int userYAdjust = (int)await userYAdjustString.EvaluateCode(engine);

					var elementLocation = ((IWebElement)element).Location;
					var seleniumWindowPosition = seleniumInstance.Manage().Window.Position;
					User32Functions.SendMouseMove(seleniumWindowPosition.X + elementLocation.X + userXAdjust, seleniumWindowPosition.Y + elementLocation.Y + userYAdjust, 
												  v_SeleniumElementAction);
					
					break;
				case "Double Left Click":
					actions.DoubleClick((IWebElement)element).Perform();
					break;

				case "Set Text":
					string textToSetString = (from rw in v_WebActionParameterTable.AsEnumerable()
										where rw.Field<string>("Parameter Name") == "Text To Set"
										select rw.Field<string>("Parameter Value")).FirstOrDefault();
					string textToSet = (string)await textToSetString.EvaluateCode(engine);

					string clearElement = (from rw in v_WebActionParameterTable.AsEnumerable()
										   where rw.Field<string>("Parameter Name") == "Clear Element Before Setting Text"
										   select rw.Field<string>("Parameter Value")).FirstOrDefault();

					if (clearElement == null)
						clearElement = "No";

					if (clearElement.ToLower() == "yes")
						((IWebElement)element).Clear();

					string[] potentialKeyPresses = textToSet.Split('[', ']');

					Type seleniumKeys = typeof(OpenQA.Selenium.Keys);
					FieldInfo[] fields = seleniumKeys.GetFields(BindingFlags.Static | BindingFlags.Public);
					string finalTextToSet = "";

					//check if chunked string contains a key press command like {ENTER}
					foreach (string chunkedString in potentialKeyPresses)
					{
						if (chunkedString == "")
							continue;

						if (fields.Any(f => f.Name.ToLower() == chunkedString.ToLower()) && textToSet.Contains("["+ chunkedString +"]"))
						{
							string keyPress = (string)fields.Where(f => f.Name.ToLower() == chunkedString.ToLower()).FirstOrDefault().GetValue(null);
							finalTextToSet += keyPress;
						}
						else
							finalTextToSet += chunkedString;
					}
					((IWebElement)element).SendKeys(finalTextToSet);
					break;

				case "Set Secure Text":
					var secureString = (from rw in v_WebActionParameterTable.AsEnumerable()
										where rw.Field<string>("Parameter Name") == "Secure String Variable"
										select rw.Field<string>("Parameter Value")).FirstOrDefault();

					string _clearElement = (from rw in v_WebActionParameterTable.AsEnumerable()
											where rw.Field<string>("Parameter Name") == "Clear Element Before Setting Text"
											select rw.Field<string>("Parameter Value")).FirstOrDefault();

					var secureStrVariable = (SecureString)await secureString.EvaluateCode(engine);
					secureString = secureStrVariable.ConvertSecureStringToString();

					if (_clearElement == null)
						_clearElement = "No";

					if (_clearElement.ToLower() == "yes")
						((IWebElement)element).Clear();

					string[] _potentialKeyPresses = secureString.Split('[', ']');

					Type _seleniumKeys = typeof(OpenQA.Selenium.Keys);
					FieldInfo[] _fields = _seleniumKeys.GetFields(BindingFlags.Static | BindingFlags.Public);
					string _finalTextToSet = "";

					//check if chunked string contains a key press command like {ENTER}
					foreach (string chunkedString in _potentialKeyPresses)
					{
						if (chunkedString == "")
							continue;

						if (_fields.Any(f => f.Name.ToLower() == chunkedString.ToLower()) && secureString.Contains("[" + chunkedString + "]"))
						{
							string keyPress = (string)_fields.Where(f => f.Name.ToLower() == chunkedString.ToLower()).FirstOrDefault().GetValue(null);
							_finalTextToSet += keyPress;
						}
						else
							_finalTextToSet += chunkedString;
					}
					((IWebElement)element).SendKeys(_finalTextToSet);
					break;

				case "Get Options":
					string applyToVarName = (from rw in v_WebActionParameterTable.AsEnumerable()
										   where rw.Field<string>("Parameter Name") == "Variable Name"
										   select rw.Field<string>("Parameter Value")).FirstOrDefault();


					string attribNameString = (from rw in v_WebActionParameterTable.AsEnumerable()
											where rw.Field<string>("Parameter Name") == "Attribute Name"
											select rw.Field<string>("Parameter Value")).FirstOrDefault();
					string attribName = (string)await attribNameString.EvaluateCode(engine);

					var optionsItems = new List<string>();
					var ele = (IWebElement)element;
					var select = new SelectElement(ele);
					var options = select.Options;

					foreach (var option in options)
					{
						var optionValue = option.GetAttribute(attribName);
						optionsItems.Add(optionValue);
					}

					optionsItems.SetVariableValue(engine, applyToVarName);
				   
					break;

				case "Select Option":
					string selectionType = (from rw in v_WebActionParameterTable.AsEnumerable()
											where rw.Field<string>("Parameter Name") == "Selection Type"
											select rw.Field<string>("Parameter Value")).FirstOrDefault();

					string selectionParamString = (from rw in v_WebActionParameterTable.AsEnumerable()
											where rw.Field<string>("Parameter Name") == "Selection Parameter"
											select rw.Field<string>("Parameter Value")).FirstOrDefault();
					string selectionParam = (string)await selectionParamString.EvaluateCode(engine);

					seleniumInstance.SwitchTo().ActiveElement();

					var el = (IWebElement)element;
					var selectionElement = new SelectElement(el);

					switch (selectionType)
					{
						case "Select By Index":
							selectionElement.SelectByIndex(int.Parse(selectionParam));
							break;
						case "Select By Text":
							selectionElement.SelectByText(selectionParam);
							break;
						case "Select By Value":
							selectionElement.SelectByValue(selectionParam);
							break;
						case "Deselect By Index":
							selectionElement.DeselectByIndex(int.Parse(selectionParam));
							break;
						case "Deselect By Text":
							selectionElement.DeselectByText(selectionParam);
							break;
						case "Deselect By Value":
							selectionElement.DeselectByValue(selectionParam);
							break;
						case "Deselect All":
							selectionElement.DeselectAll();
							break;
						default:
							throw new NotImplementedException();
					}
					break;

				case "Get Text":
				case "Get Attribute":
				case "Get Count":
					string VariableName = (from rw in v_WebActionParameterTable.AsEnumerable()
										   where rw.Field<string>("Parameter Name") == "Variable Name"
										   select rw.Field<string>("Parameter Value")).FirstOrDefault();

					string attributeNameString = (from rw in v_WebActionParameterTable.AsEnumerable()
											where rw.Field<string>("Parameter Name") == "Attribute Name"
											select rw.Field<string>("Parameter Value")).FirstOrDefault();
					string attributeName = (string)await attributeNameString.EvaluateCode(engine);

					string elementValue;
					if (v_SeleniumElementAction == "Get Text")
						elementValue = element.Text;
					else if (v_SeleniumElementAction == "Get Count")
					{
						elementValue = "1";
						if (element is ReadOnlyCollection<IWebElement>)
							elementValue = ((ReadOnlyCollection<IWebElement>)element).Count().ToString();
					}
					else
						elementValue = ((IWebElement)element).GetAttribute(attributeName);

					elementValue.SetVariableValue(engine, VariableName);
					break;

				case "Get Matching Element(s)":
					var variableName = (from rw in v_WebActionParameterTable.AsEnumerable()
										where rw.Field<string>("Parameter Name") == "Variable Name"
										select rw.Field<string>("Parameter Value")).FirstOrDefault();

					if (!(element is IWebElement))
					{
						//create element list
						List<IWebElement> elementList = new List<IWebElement>();
						foreach (IWebElement item in element)
						{
							elementList.Add(item);
						}
						elementList.SetVariableValue(engine, variableName);
					}
					else
						((IWebElement)element).SetVariableValue(engine, variableName);                    
					break;

				case "Get Table":
					var DTVariableName = (from rw in v_WebActionParameterTable.AsEnumerable()
										  where rw.Field<string>("Parameter Name") == "Variable Name"
										  select rw.Field<string>("Parameter Value")).FirstOrDefault();

					// Get HTML (Source) of the Element
					string tableHTML = ((IWebElement)element).GetAttribute("innerHTML").ToString();
					HtmlDocument doc = new HtmlDocument();

					//Load Source (String) as HTML Document
					doc.LoadHtml(tableHTML);

					//Get Header Tags
					var headers = doc.DocumentNode.SelectNodes("//tr/th");
					DataTable DT = new DataTable();

					//If headers found
					if (headers != null && headers.Count != 0)
					{
						// add columns from th (headers)
						foreach (HtmlNode header in headers)
							DT.Columns.Add(Regex.Replace(header.InnerText, @"\t|\n|\r", "").Trim()); 
					}
					else
					{
						var columnsCount = doc.DocumentNode.SelectSingleNode("//tr[1]").ChildNodes.Where(node=>node.Name=="td").Count();
						DT.Columns.AddRange((Enumerable.Range(1, columnsCount).Select(dc => new DataColumn())).ToArray());
					}

					// select rows with td elements and load each row (containing <td> tags) into DataTable
					foreach (var row in doc.DocumentNode.SelectNodes("//tr[td]"))
						DT.Rows.Add(row.SelectNodes("td").Select(td => Regex.Replace(td.InnerText, @"\t|\n|\r", "").Trim()).ToArray());

					DT.SetVariableValue(engine, DTVariableName);
					break;

				case "Clear Element":
					((IWebElement)element).Clear();
					break;

				case "Switch to Frame":
					if (seleniumSearchParamRows.Count == 0)
						seleniumInstance.SwitchTo().DefaultContent();
					else
						seleniumInstance.SwitchTo().Frame((IWebElement)element);
					break;
				case "Wait For Element To Exist":
					break;
				case "Element Exists":
					string existsBoolVariableName = (from rw in v_WebActionParameterTable.AsEnumerable()
													where rw.Field<string>("Parameter Name") == "Variable Name"
													select rw.Field<string>("Parameter Value")).FirstOrDefault();

					if (element == null)
						false.SetVariableValue(engine, existsBoolVariableName);
					else
						true.SetVariableValue(engine, existsBoolVariableName);

					break;
				default:
					throw new Exception("Element Action was not found");
			}
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			//create helper control
			CommandItemControl helperControl = new CommandItemControl();
			helperControl.Padding = new Padding(10, 0, 0, 0);
			helperControl.ForeColor = Color.AliceBlue;
			helperControl.Font = new Font("Segoe UI Semilight", 10);
			helperControl.CommandImage = Resources.command_camera;
			helperControl.CommandDisplay = "Element Recorder";
			helperControl.Click += new EventHandler((s, e) => ShowRecorder(s, e, editor, commandControls));

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));

			if (v_SeleniumSearchParameters.Rows.Count == 0)
			{
				v_SeleniumSearchParameters.Rows.Add(true, "\"XPath\"", "");
				v_SeleniumSearchParameters.Rows.Add(false, "\"ID\"", "");
				v_SeleniumSearchParameters.Rows.Add(false, "\"Name\"", "");
				v_SeleniumSearchParameters.Rows.Add(false, "\"Tag Name\"", "");
				v_SeleniumSearchParameters.Rows.Add(false, "\"Class Name\"", "");
				v_SeleniumSearchParameters.Rows.Add(false, "\"Link Text\"", "");
				v_SeleniumSearchParameters.Rows.Add(false, "\"CSS Selector\"", "");
			}
			//create search parameters   
			RenderedControls.Add(commandControls.CreateDefaultLabelFor("v_SeleniumSearchParameters", this));
			RenderedControls.Add(helperControl);

			//create search param grid
			_searchParametersGridViewHelper = commandControls.CreateDefaultDataGridViewFor("v_SeleniumSearchParameters", this);
			_searchParametersGridViewHelper.MouseEnter += ActionParametersGridViewHelper_MouseEnter;

			DataGridViewCheckBoxColumn enabled = new DataGridViewCheckBoxColumn();
			enabled.HeaderText = "Enabled";
			enabled.DataPropertyName = "Enabled";
			enabled.FillWeight = 30;
			_searchParametersGridViewHelper.Columns.Add(enabled);

			DataGridViewTextBoxColumn propertyName = new DataGridViewTextBoxColumn();
			propertyName.HeaderText = "Parameter Name";
			propertyName.DataPropertyName = "Parameter Name";
			propertyName.FillWeight = 40;
			_searchParametersGridViewHelper.Columns.Add(propertyName);

			DataGridViewTextBoxColumn propertyValue = new DataGridViewTextBoxColumn();
			propertyValue.HeaderText = "Parameter Value";
			propertyValue.DataPropertyName = "Parameter Value";
			_searchParametersGridViewHelper.Columns.Add(propertyValue);

			RenderedControls.AddRange(commandControls.CreateUIHelpersFor("v_SeleniumSearchParameters", this, new Control[] { _searchParametersGridViewHelper }, editor));
			RenderedControls.Add(_searchParametersGridViewHelper);

			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_SeleniumSearchOption", this, editor));

			_elementActionDropdown = commandControls.CreateDropdownFor("v_SeleniumElementAction", this);
			RenderedControls.Add(commandControls.CreateDefaultLabelFor("v_SeleniumElementAction", this));
			RenderedControls.AddRange(commandControls.CreateUIHelpersFor("v_SeleniumElementAction", this, new Control[] { _elementActionDropdown }, editor));
			_elementActionDropdown.SelectionChangeCommitted += SeleniumAction_SelectionChangeCommitted;
			RenderedControls.Add(_elementActionDropdown);

			_actionParametersControls = new List<Control>();
			_actionParametersControls.Add(commandControls.CreateDefaultLabelFor("v_WebActionParameterTable", this));

			_actionParametersGridViewHelper = commandControls.CreateDefaultDataGridViewFor("v_WebActionParameterTable", this);
			_actionParametersGridViewHelper.AllowUserToAddRows = false;
			_actionParametersGridViewHelper.AllowUserToDeleteRows = false;
			_actionParametersGridViewHelper.MouseEnter += ActionParametersGridViewHelper_MouseEnter;

			_actionParametersControls.AddRange(commandControls.CreateUIHelpersFor("v_WebActionParameterTable", this, new Control[] { _actionParametersGridViewHelper }, editor));
			_actionParametersControls.Add(_actionParametersGridViewHelper);
			RenderedControls.AddRange(_actionParametersControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_Timeout", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			string searchParameterName = (from rw in v_SeleniumSearchParameters.AsEnumerable()
										  where rw.Field<string>("Enabled") == "True"
										  select rw.Field<string>("Parameter Name")).FirstOrDefault();

			string searchParameterValue = (from rw in v_SeleniumSearchParameters.AsEnumerable()
										  where rw.Field<string>("Enabled") == "True"
										  select rw.Field<string>("Parameter Value")).FirstOrDefault();


			return base.GetDisplayValue() + $" [{v_SeleniumElementAction} - {v_SeleniumSearchOption} by {searchParameterName}" + 
											$" '{searchParameterValue}' - Instance Name '{v_InstanceName}']";
		}

		private void ActionParametersGridViewHelper_MouseEnter(object sender, EventArgs e)
		{
			SeleniumAction_SelectionChangeCommitted(null, null);
		}

		public void ShowRecorder(object sender, EventArgs e, IfrmCommandEditor editor, ICommandControls commandControls)
		{
			//create recorder
			IfrmWebElementRecorder newElementRecorder = commandControls.CreateWebElementRecorderForm(editor.HTMLElementRecorderURL);
			newElementRecorder.ScriptContext = editor.ScriptContext;
			newElementRecorder.CheckBox_StopOnClick(true);
			//show form
			((Form)newElementRecorder).ShowDialog();

			editor.HTMLElementRecorderURL = newElementRecorder.StartURL;

			try
			{
				if (newElementRecorder.SearchParameters != null)
				{
					v_SeleniumSearchParameters.Rows.Clear();

					foreach (DataRow rw in newElementRecorder.SearchParameters.Rows)
						v_SeleniumSearchParameters.ImportRow(rw);

					_searchParametersGridViewHelper.DataSource = v_SeleniumSearchParameters;
					_searchParametersGridViewHelper.Refresh();
				}               
			}
			catch (Exception)
			{
				//Search parameter not found
			}
		}

		public void SeleniumAction_SelectionChangeCommitted(object sender, EventArgs e)
		{
			SeleniumElementActionCommand cmd = this;
			DataTable actionParameters = cmd.v_WebActionParameterTable;

			if (sender != null)
				actionParameters.Rows.Clear();

			switch (_elementActionDropdown.SelectedItem)
			{
				case "Invoke Click":
				case "Left Click":
				case "Right Click":
				case "Double Left Click":
				case "Clear Element":
					foreach (var ctrl in _actionParametersControls)
						ctrl.Hide();
					break;  
					
				case "Middle Click":               
					foreach (var ctrl in _actionParametersControls)
						ctrl.Show();

					if (sender != null)
					{
						actionParameters.Rows.Add("X Adjustment", 0);
						actionParameters.Rows.Add("Y Adjustment", 0);
					}
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
				case "Get Matching Element(s)":
				case "Get Table":
				case "Get Count":
					foreach (var ctrl in _actionParametersControls)
						ctrl.Show();

					if (sender != null)
						actionParameters.Rows.Add("Variable Name");
					break;

				case "Get Attribute":
					foreach (var ctrl in _actionParametersControls)
						ctrl.Show();

					if (sender != null)
					{
						actionParameters.Rows.Add("Attribute Name");
						actionParameters.Rows.Add("Variable Name");
					}
					break;

				case "Get Options":
					foreach (var ctrl in _actionParametersControls)
						ctrl.Show();

					if (sender != null)
					{
						actionParameters.Rows.Add("Attribute Name");
						actionParameters.Rows.Add("Variable Name");
					}
					break;

				case "Select Option":
					foreach (var ctrl in _actionParametersControls)
						ctrl.Show();

					if (sender != null)
					{
						actionParameters.Rows.Add("Selection Type");
						actionParameters.Rows.Add("Selection Parameter");
					}

					DataGridViewComboBoxCell selectionTypeBox = new DataGridViewComboBoxCell();
					selectionTypeBox.Items.Add("Select By Index");
					selectionTypeBox.Items.Add("Select By Text");
					selectionTypeBox.Items.Add("Select By Value");
					selectionTypeBox.Items.Add("Deselect By Index");
					selectionTypeBox.Items.Add("Deselect By Text");
					selectionTypeBox.Items.Add("Deselect By Value");
					selectionTypeBox.Items.Add("Deselect All");

					//assign cell as a combobox
					if (sender != null)
						_actionParametersGridViewHelper.Rows[0].Cells[1].Value = "Select By Text";

					if (_actionParametersGridViewHelper.Rows.Count > 0)
						_actionParametersGridViewHelper.Rows[0].Cells[1] = selectionTypeBox;
					break;

				case "Wait For Element To Exist":
					foreach (var ctrl in _actionParametersControls)
						ctrl.Hide();
					break;

				case "Element Exists":
					foreach (var ctrl in _actionParametersControls)
						ctrl.Show();

					if (sender != null)
						actionParameters.Rows.Add("Variable Name");
					break;

				case "Switch to frame":
					foreach (var ctrl in _actionParametersControls)
						ctrl.Hide();
					break;

				default:
					break;
			}

			_actionParametersGridViewHelper.Columns[0].ReadOnly = true;
			_actionParametersGridViewHelper.DataSource = v_WebActionParameterTable;
		}
	}
}
