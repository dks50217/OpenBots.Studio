﻿using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Utilities.CommonUtilities;

using SimpleNLG;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Windows.Forms;

namespace OpenBots.Commands.NLG
{
	[Serializable]
	[Category("NLG Commands")]
	[Description("This command creates a Natural Language Generation Instance.")]
	public class CreateNLGInstanceCommand : ScriptCommand
	{

		[Required]
		[DisplayName("NLG Instance Name")]
		[Description("Enter a unique name that will represent the application instance.")]
		[SampleUsage("MyNLGInstance")]
		[Remarks("This unique name allows you to refer to the instance by name in future commands, " +
				 "ensuring that the commands you specify run against the correct application.")]
		public string v_InstanceName { get; set; }

		public CreateNLGInstanceCommand()
		{
			CommandName = "CreateNLGInstanceCommand";
			SelectionName = "Create NLG Instance";
			CommandEnabled = true;
			
			v_InstanceName = "DefaultNLG";
		}

		public override void RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
  
			Lexicon lexicon = Lexicon.getDefaultLexicon();
			NLGFactory nlgFactory = new NLGFactory(lexicon);
			SPhraseSpec p = nlgFactory.createClause();

			p.AddAppInstance(engine, v_InstanceName);
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
			
			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [New Instance Name '{v_InstanceName}']";
		}
	}
}