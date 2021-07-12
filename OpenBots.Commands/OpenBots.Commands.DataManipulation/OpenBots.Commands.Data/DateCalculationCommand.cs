﻿using Newtonsoft.Json;
using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Interfaces;
using OpenBots.Core.Properties;
using OpenBots.Core.Utilities.CommonUtilities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenBots.Commands.Data
{
	[Serializable]
	[Category("Data Commands")]
	[Description("This command performs a specific operation on a date and saves the result in a variable.")]
	public class DateCalculationCommand : ScriptCommand
	{
		[Required]
		[DisplayName("Date")]
		[Description("Specify either text or a variable that contains the date.")]
		[SampleUsage("new DateTime(2020, 2, 20) || vDate || DateTime.Now")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(DateTime) })]
		public string v_InputDate { get; set; }

		[Required]
		[DisplayName("Calculation Method")]
		[PropertyUISelectionOption("Add Second(s)")]
		[PropertyUISelectionOption("Add Minute(s)")]
		[PropertyUISelectionOption("Add Hour(s)")]
		[PropertyUISelectionOption("Add Day(s)")]
		[PropertyUISelectionOption("Add Month(s)")]
		[PropertyUISelectionOption("Add Year(s)")]
		[PropertyUISelectionOption("Subtract Second(s)")]
		[PropertyUISelectionOption("Subtract Minute(s)")]
		[PropertyUISelectionOption("Subtract Hour(s)")]
		[PropertyUISelectionOption("Subtract Day(s)")]
		[PropertyUISelectionOption("Subtract Month(s)")]
		[PropertyUISelectionOption("Subtract Year(s)")]
		[PropertyUISelectionOption("Get Next Day")]
		[PropertyUISelectionOption("Get Next Month")]
		[PropertyUISelectionOption("Get Next Year")]
		[PropertyUISelectionOption("Get Previous Day")]
		[PropertyUISelectionOption("Get Previous Month")]
		[PropertyUISelectionOption("Get Previous Year")]
		[Description("Select the date operation.")]
		[SampleUsage("")]
		[Remarks("The selected operation will be applied to the input date value and result will be stored in the output variable.")]
		public string v_CalculationMethod { get; set; }

		[Required]
		[DisplayName("Increment Value")]
		[Description("Specify how many units to increment by.")]
		[SampleUsage("15 || vIncrement")]
		[Remarks("You can use negative numbers which will do the opposite, ex. Subtract Days and an increment of -5 will Add Days.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(double) })]
		public string v_Increment { get; set; }

		[DisplayName("Date Format (Optional)")]
		[Description("Specify the output date format.")]
		[SampleUsage("\"MM/dd/yy hh:mm:ss\" || \"MM/dd/yyyy\" || vDateFormat")]
		[Remarks("You can specify either a valid DateTime, Date or Time Format; an invalid format will result in an error.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_ToStringFormat { get; set; }

		[Required]
		[Editable(false)]
		[DisplayName("Output Date Variable")]
		[Description("Create a new variable or select a variable from the list.")]
		[SampleUsage("vUserVariable")]
		[Remarks("New variables/arguments may be instantiated by utilizing the Ctrl+K/Ctrl+J shortcuts.")]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_OutputUserVariableName { get; set; }

		[JsonIgnore]
		[Browsable(false)]
		private List<Control> _stringFormatControls;

		[JsonIgnore]
		[Browsable(false)]
		private bool _hasRendered;

		public DateCalculationCommand()
		{
			CommandName = "DateCalculationCommand";
			SelectionName = "Date Calculation";
			CommandEnabled = true;
			CommandIcon = Resources.command_stopwatch;

			v_InputDate = "DateTime.Now";
			v_CalculationMethod = "Add Second(s)";
			v_ToStringFormat = "\"MM/dd/yyyy hh:mm:ss\"";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;

			var formatting = (string)await v_ToStringFormat.EvaluateCode(engine);
			var variableIncrement = Convert.ToDouble(await v_Increment.EvaluateCode(engine));
			DateTime variableDate = (DateTime)await v_InputDate.EvaluateCode(engine);

			//get increment value
			double requiredInterval = variableIncrement;

			dynamic dateTimeValue;

			//perform operation
			switch (v_CalculationMethod)
			{
				case "Add Second(s)":
					dateTimeValue = variableDate.AddSeconds(requiredInterval);
					break;
				case "Add Minute(s)":
					dateTimeValue = variableDate.AddMinutes(requiredInterval);
					break;
				case "Add Hour(s)":
					dateTimeValue = variableDate.AddHours(requiredInterval);
					break;
				case "Add Day(s)":
					dateTimeValue = variableDate.AddDays(requiredInterval);
					break;
				case "Add Month(s)":
					dateTimeValue = variableDate.AddMonths((int)requiredInterval);
					break;
				case "Add Year(s)":
					dateTimeValue = variableDate.AddYears((int)requiredInterval);
					break;
				case "Subtract Second(s)":
					dateTimeValue = variableDate.AddSeconds(requiredInterval * -1);
					break;
				case "Subtract Minute(s)":
					dateTimeValue = variableDate.AddMinutes(requiredInterval * -1);
					break;
				case "Subtract Hour(s)":
					dateTimeValue = variableDate.AddHours(requiredInterval * -1);
					break;
				case "Subtract Day(s)":
					dateTimeValue = variableDate.AddDays(requiredInterval * -1);
					break;
				case "Subtract Month(s)":
					dateTimeValue = variableDate.AddMonths((int)requiredInterval * -1);
					break;
				case "Subtract Year(s)":
					dateTimeValue = variableDate.AddYears((int)requiredInterval * -1);
					break;
				case "Get Next Day":
					dateTimeValue = variableDate.AddDays(requiredInterval).Day;
					break;
				case "Get Next Month":
					dateTimeValue = variableDate.AddMonths((int)requiredInterval).Month;
					break;
				case "Get Next Year":
					dateTimeValue = variableDate.AddYears((int)requiredInterval).Year;
					break;
				case "Get Previous Day":
					dateTimeValue = variableDate.AddDays(requiredInterval * -1).Day;
					break;
				case "Get Previous Month":
					dateTimeValue = variableDate.AddMonths((int)requiredInterval * -1).Month;
					break;
				case "Get Previous Year":
					dateTimeValue = variableDate.AddYears((int)requiredInterval * -1).Year;
					break;
				default:
					dateTimeValue = "";
					break;
			}

			string stringDateFormatted;

			//handle if formatter is required
			if (!string.IsNullOrEmpty(formatting.Trim()))
				stringDateFormatted = ((DateTime)dateTimeValue).ToString(formatting);
			else
				stringDateFormatted = ((object)dateTimeValue).ToString();

			//store string (Result) in variable
			stringDateFormatted.SetVariableValue(engine, v_OutputUserVariableName);
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			//create standard group controls
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InputDate", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_CalculationMethod", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_Increment", this, editor));

			((ComboBox)RenderedControls[4]).SelectedIndexChanged += calculationMethodComboBox_SelectedIndexChanged;

			_stringFormatControls = new List<Control>();
			_stringFormatControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_ToStringFormat", this, editor));

			RenderedControls.AddRange(_stringFormatControls);

			RenderedControls.AddRange(commandControls.CreateDefaultOutputGroupFor("v_OutputUserVariableName", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			string operand = "Add";
			string interval = "";

			//determine operand and interval
			if (!string.IsNullOrEmpty(v_CalculationMethod))
			{
				operand = v_CalculationMethod.Split(' ').First();
				interval = v_CalculationMethod.Split(' ').Last();
			}
			
			//additional language handling based on selection made
			string operandLanguage;
			if (operand == "Add")
				operandLanguage = "to";
			else
				operandLanguage = "from";

			if (operand == "Get")
				operand = v_CalculationMethod.Replace(interval, "").TrimEnd();

			//return value
			return base.GetDisplayValue() + $" [{operand} '{v_Increment}' {interval} {operandLanguage} '{v_InputDate}' - Store Date in '{v_OutputUserVariableName}']";
		}

		public override void Shown()
		{
			base.Shown();
			_hasRendered = true;
			calculationMethodComboBox_SelectedIndexChanged(null, null);
		}

		private void calculationMethodComboBox_SelectedIndexChanged(object sender, EventArgs e)
		{
			if (!((ComboBox)RenderedControls[4]).Text.StartsWith("Get") && _hasRendered)
			{
				foreach (var ctrl in _stringFormatControls)
				{
					ctrl.Visible = true;
					if (ctrl is TextBox && string.IsNullOrEmpty(((TextBox)ctrl).Text))
						((TextBox)ctrl).Text = "\"MM/dd/yyyy hh:mm:ss\"";
				}                                    
			}
			else if(_hasRendered)
			{
				foreach (var ctrl in _stringFormatControls)
				{
					ctrl.Visible = false;
					if (ctrl is TextBox)
						((TextBox)ctrl).Clear();
				}
			}
		}
	}
}