﻿using Newtonsoft.Json;
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
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using OBDataTable = System.Data.DataTable;

namespace OpenBots.Commands.Data
{
    [Serializable]
	[Category("Data Commands")]
	[Description("This command performs advanced text extraction.")]
	public class TextExtractionCommand : ScriptCommand
	{
		[Required]
		[DisplayName("Text Data")]
		[Description("Provide a variable or text value.")]
		[SampleUsage("\"Hello, welcome to OpenBots\" || vTextData")]
		[Remarks("Providing data of a type other than a 'String' will result in an error.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_InputText { get; set; }

		[Required]
		[DisplayName("Text Extraction Type")]
		[PropertyUISelectionOption("Extract All After Text")]
		[PropertyUISelectionOption("Extract All Before Text")]
		[PropertyUISelectionOption("Extract All Between Text")]
		[Description("Select the type of extraction.")]
		[SampleUsage("")]
		[Remarks("For trailing text, use 'After Text'. For leading text, use 'Before Text'. For text between two substrings, use 'Between Text'.")]
		public string v_TextExtractionType { get; set; }

		[Required]
		[DisplayName("Extraction Parameters")]
		[Description("Define the required extraction parameters, which is dependent on the type of extraction.")]
		[SampleUsage("[\"Welcome\", 0] || [vSubstring, vOccurences]")]
		[Remarks("Set parameter values for each parameter name based on the extraction type.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string), typeof(int) })]
		public OBDataTable v_TextExtractionTable { get; set; }

		[Required]
		[Editable(false)]
		[DisplayName("Output Text Variable")]
		[Description("Create a new variable or select a variable from the list.")]
		[SampleUsage("vUserVariable")]
		[Remarks("New variables/arguments may be instantiated by utilizing the Ctrl+K/Ctrl+J shortcuts.")]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_OutputUserVariableName { get; set; }

		[JsonIgnore]
		[Browsable(false)]
		private DataGridView _textExtractionGridViewHelper;

		public TextExtractionCommand()
		{
			CommandName = "TextExtractionCommand";
			SelectionName = "Text Extraction";
			CommandEnabled = true;
			CommandIcon = Resources.command_function;

			//define parameter table
			v_TextExtractionTable = new OBDataTable
			{
				TableName = DateTime.Now.ToString("TextExtractorParamTable" + DateTime.Now.ToString("MMddyy.hhmmss"))
			};

			v_TextExtractionTable.Columns.Add("Parameter Name");
			v_TextExtractionTable.Columns.Add("Parameter Value");
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var variableInput = (string)await v_InputText.EvaluateCode(engine);

			string variableLeading, variableTrailing, extractedText;
			int skipOccurences;

			//handle extraction cases
			switch (v_TextExtractionType)
			{
				case "Extract All After Text":
					//extract trailing texts            
					variableLeading = (string)await GetParameterValue("Leading Text").EvaluateCode(engine);
					skipOccurences = (int)await GetParameterValue("Skip Past Occurences").EvaluateCode(engine);
					extractedText = ExtractLeadingText(variableInput, variableLeading, skipOccurences);
					break;
				case "Extract All Before Text":
					//extract leading text
					variableTrailing = (string)await GetParameterValue("Trailing Text").EvaluateCode(engine);
					skipOccurences = (int)await GetParameterValue("Skip Past Occurences").EvaluateCode(engine);
					extractedText = ExtractTrailingText(variableInput, variableTrailing, skipOccurences);
					break;
				case "Extract All Between Text":
					//extract leading and then trailing which gives the items between
					variableLeading = (string)await GetParameterValue("Leading Text").EvaluateCode(engine);
					variableTrailing = (string)await GetParameterValue("Trailing Text").EvaluateCode(engine);
					skipOccurences = (int)await GetParameterValue("Skip Past Occurences").EvaluateCode(engine);

					//extract leading
					extractedText = ExtractLeadingText(variableInput, variableLeading, skipOccurences);

					//extract trailing -- assume we will take to the first item
					extractedText = ExtractTrailingText(extractedText, variableTrailing, 0);

					break;
				default:
					throw new NotImplementedException("Extraction Type Not Implemented: " + v_TextExtractionType);
			}

			extractedText.SetVariableValue(engine, v_OutputUserVariableName);
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InputText", this, editor));

			RenderedControls.Add(commandControls.CreateDefaultLabelFor("v_TextExtractionType", this));
			var selectionControl = commandControls.CreateDropdownFor("v_TextExtractionType", this);
			RenderedControls.AddRange(commandControls.CreateUIHelpersFor("v_TextExtractionType", this, new Control[] { selectionControl }, editor));
			selectionControl.SelectionChangeCommitted += TextExtraction_SelectionChangeCommitted;
			RenderedControls.Add(selectionControl);

			var textExtractionGridViewControls = new List<Control>();
			textExtractionGridViewControls.Add(commandControls.CreateDefaultLabelFor("v_TextExtractionTable", this));

			_textExtractionGridViewHelper = commandControls.CreateDefaultDataGridViewFor("v_TextExtractionTable", this);
			if (_textExtractionGridViewHelper.Columns.Count > 0)
				_textExtractionGridViewHelper.Columns[0].ReadOnly = true;
			_textExtractionGridViewHelper.AllowUserToAddRows = false;
			_textExtractionGridViewHelper.AllowUserToDeleteRows = false;
			_textExtractionGridViewHelper.MouseEnter += TextExtractionGridViewHelper_MouseEnter;

			textExtractionGridViewControls.AddRange(commandControls.CreateUIHelpersFor("v_TextExtractionTable", this, new Control[] { _textExtractionGridViewHelper }, editor));
			textExtractionGridViewControls.Add(_textExtractionGridViewHelper);
			RenderedControls.AddRange(textExtractionGridViewControls);

			RenderedControls.AddRange(commandControls.CreateDefaultOutputGroupFor("v_OutputUserVariableName", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Extract Text From '{v_InputText}' - Store Text in '{v_OutputUserVariableName}']";
		}

		private void TextExtraction_SelectionChangeCommitted(object sender, EventArgs e)
		{
			ComboBox extractionAction = (ComboBox)sender;
			TextExtractionCommand cmd = this;
			OBDataTable textExtractionTable = cmd.v_TextExtractionTable;

			if (sender != null)
            {
				textExtractionTable.Rows.Clear();

				switch (extractionAction.SelectedItem)
				{
					case "Extract All After Text":
						textExtractionTable.Rows.Add("Leading Text", "");
						textExtractionTable.Rows.Add("Skip Past Occurences", "0");
						break;
					case "Extract All Before Text":
						textExtractionTable.Rows.Add("Trailing Text", "");
						textExtractionTable.Rows.Add("Skip Past Occurences", "0");
						break;
					case "Extract All Between Text":
						textExtractionTable.Rows.Add("Leading Text", "");
						textExtractionTable.Rows.Add("Trailing Text", "");
						textExtractionTable.Rows.Add("Skip Past Occurences", "0");
						break;
					default:
						break;
				}
			}

			_textExtractionGridViewHelper.Columns[0].ReadOnly = true;
			_textExtractionGridViewHelper.DataSource = v_TextExtractionTable;
		}

		private void TextExtractionGridViewHelper_MouseEnter(object sender, EventArgs e)
		{
			TextExtraction_SelectionChangeCommitted(null, null);
		}

		private string GetParameterValue(string parameterName)
		{
			return (from rw in v_TextExtractionTable.AsEnumerable()
					 where rw.Field<string>("Parameter Name") == parameterName
					 select rw.Field<string>("Parameter Value")).FirstOrDefault();
		}

		private string ExtractLeadingText(string input, string substring, int occurences)
		{
			//verify the occurence index
			int leadingOccurenceIndex = occurences;

			//find index matches
			var leadingOccurencesFound = Regex.Matches(input, substring).Cast<Match>().Select(m => m.Index).ToList();

			//handle if we are searching beyond what was found
			if (leadingOccurenceIndex >= leadingOccurencesFound.Count)
			{
				throw new Exception("No value was found after skipping " + leadingOccurenceIndex + " instance(s).  Only " + 
					leadingOccurencesFound.Count + " instances exist.");
			}

			//declare start position
			var startPosition = leadingOccurencesFound[leadingOccurenceIndex] + substring.Length;

			//substring and apply to variable
			return input.Substring(startPosition);
		}

		private string ExtractTrailingText(string input, string substring, int occurences)
		{
			//verify the occurence index
			int leadingOccurenceIndex = occurences;

			//find index matches
			var trailingOccurencesFound = Regex.Matches(input, substring).Cast<Match>().Select(m => m.Index).ToList();

			//handle if we are searching beyond what was found
			if (leadingOccurenceIndex >= trailingOccurencesFound.Count)
			{
				throw new Exception("No value was found after skipping " + leadingOccurenceIndex + " instance(s).  Only " + 
					trailingOccurencesFound.Count + " instances exist.");
			}

			//declare start position
			var endPosition = trailingOccurencesFound[leadingOccurenceIndex];

			//substring
			return input.Substring(0, endPosition);
		}
	}
}