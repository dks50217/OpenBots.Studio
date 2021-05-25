﻿using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Properties;
using OpenBots.Core.Utilities.CommonUtilities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenBots.Commands.Data
{
	[Serializable]
	[Category("Data Commands")]
	[Description("This command replaces an existing substring in a string and saves the result in a variable.")]
	public class ReplaceTextCommand : ScriptCommand
	{
		[Required]
		[DisplayName("Text Data")]
		[Description("Provide a variable or text value.")]
		[SampleUsage("\"Hello John\" || vTextData")]
		[Remarks("Providing data of a type other than a 'String' will result in an error.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_InputText { get; set; }

		[Required]
		[DisplayName("Old Text")]
		[Description("Specify the old value of the text that will be replaced.")]
		[SampleUsage("\"Hello\" || vOldText")]
		[Remarks("'Hello' in 'Hello John' would be targeted for replacement.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_OldText { get; set; }

		[Required]
		[DisplayName("New Text")]
		[Description("Specify the new value to replace the old value.")]
		[SampleUsage("\"Hi\" || vNewText")]
		[Remarks("'Hi' would be replaced with 'Hello' to form 'Hi John'.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_NewText { get; set; }

		[Required]
		[Editable(false)]
		[DisplayName("Output Text Variable")]
		[Description("Create a new variable or select a variable from the list.")]
		[SampleUsage("vUserVariable")]
		[Remarks("New variables/arguments may be instantiated by utilizing the Ctrl+K/Ctrl+J shortcuts.")]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_OutputUserVariableName { get; set; }

		public ReplaceTextCommand()
		{
			CommandName = "ReplaceTextCommand";
			SelectionName = "Replace Text";
			CommandEnabled = true;
			CommandIcon = Resources.command_string;
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			string replacementVariable = (string)await v_InputText.EvaluateCode(engine);
			string replacementText = (string)await v_OldText.EvaluateCode(engine);
			string replacementValue = (string)await v_NewText.EvaluateCode(engine);

			replacementVariable = replacementVariable.Replace(replacementText, replacementValue);

			replacementVariable.SetVariableValue(engine, v_OutputUserVariableName);
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InputText", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_OldText", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_NewText", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultOutputGroupFor("v_OutputUserVariableName", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Replace '{v_OldText}' With '{v_NewText}'- Store Text in '{v_OutputUserVariableName}']";
		}
	}
}