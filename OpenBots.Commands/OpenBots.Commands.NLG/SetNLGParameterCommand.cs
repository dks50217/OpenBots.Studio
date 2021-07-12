﻿using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Interfaces;
using OpenBots.Core.Model.ApplicationModel;
using OpenBots.Core.Properties;
using OpenBots.Core.Utilities.CommonUtilities;
using SimpleNLG;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenBots.Commands.NLG
{
	[Serializable]
	[Category("NLG Commands")]
	[Description("This command defines a Natural Language Generation parameter.")]
	public class SetNLGParameterCommand : ScriptCommand
	{

		[Required]
		[DisplayName("NLG Instance Name")]
		[Description("Enter the unique instance that was specified in the **Create NLG Instance** command.")]
		[SampleUsage("MyNLGInstance")]
		[Remarks("Failure to enter the correct instance name or failure to first call the **Create NLG Instance** command will cause an error.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(OBAppInstance) })]
		public string v_InstanceName { get; set; }

		[Required]
		[DisplayName("NLG Parameter Type")]
		[PropertyUISelectionOption("Set Subject")]
		[PropertyUISelectionOption("Set Verb")]
		[PropertyUISelectionOption("Set Object")]
		[PropertyUISelectionOption("Add Complement")]
		[PropertyUISelectionOption("Add Modifier")]
		[PropertyUISelectionOption("Add Pre-Modifier")]
		[PropertyUISelectionOption("Add Front Modifier")]
		[PropertyUISelectionOption("Add Post Modifier")]
		[Description("Select the appropriate Natural Language Generation Parameter.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_ParameterType { get; set; }

		[Required]
		[DisplayName("Input Value")]
		[Description("Enter the value that should be associated with the parameter")]
		[SampleUsage("\"Hello\" || vValue")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_Parameter { get; set; }

		public SetNLGParameterCommand()
		{
			CommandName = "SetNLGParameterCommand";
			SelectionName = "Set NLG Parameter";
			CommandEnabled = true;
			CommandIcon = Resources.command_nlg;

			v_ParameterType = "Set Subject";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var p = (SPhraseSpec)((OBAppInstance)await v_InstanceName.EvaluateCode(engine)).Value;

			var userInput = (string)await v_Parameter.EvaluateCode(engine);

			switch (v_ParameterType)
			{
				case "Set Subject":
					p.setSubject(userInput);
					break;
				case "Set Object":
					p.setObject(userInput);
					break;
				case "Set Verb":
					p.setVerb(userInput);
					break;
				case "Add Complement":
					p.addComplement(userInput);
					break;
				case "Add Modifier":
					p.addModifier(userInput);             
					break;
				case "Add Front Modifier":
					p.addFrontModifier(userInput);
					break;
				case "Add Post Modifier":
					p.addPostModifier(userInput);
					break;
				case "Add Pre-Modifier":
					p.addPreModifier(userInput);
					break;
				default:
					break;
			}

			//add to app instance to track
			new OBAppInstance(v_InstanceName, p).SetVariableValue(engine, v_InstanceName);
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_ParameterType", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_Parameter", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [{v_ParameterType} '{v_Parameter}' - Instance Name '{v_InstanceName}']";
		}
	}
}
