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

namespace OpenBots.Commands.WebBrowser
{
	[Serializable]
	[Category("Web Browser Commands")]
	[Description("This command parses and extracts data from an HTML source object or a successful **GetHTMLSourceCommand**.")]

	public class QueryHTMLSourceCommand : ScriptCommand
	{
		[Required]
		[DisplayName("HTML")]
		[Description("Enter the HTML to be queried.")]
		[SampleUsage("<!DOCTYPE html><html><head><title>Example</title></head></html> || vMyHTML")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_HTMLVariable { get; set; }

		[Required]
		[DisplayName("XPath Query")]
		[Description("Enter the XPath Query and the item will be extracted.")]
		[SampleUsage("\"@//*[@id=\\\"aso_search_form_anchor\\\"]/div/input\" || vMyXPath")]
		[Remarks("You can use Chrome Dev Tools to click an element and copy the XPath.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_XPathQuery { get; set; }

		[Required]
		[Editable(false)]
		[DisplayName("Output Query Result Variable")]
		[Description("Create a new variable or select a variable from the list.")]
		[SampleUsage("vUserVariable")]
		[Remarks("New variables/arguments may be instantiated by utilizing the Ctrl+K/Ctrl+J shortcuts.")]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_OutputUserVariableName { get; set; }

		public QueryHTMLSourceCommand()
		{
			CommandName = "QueryHTMLSourceCommand";
			SelectionName = "Query HTML Source";
			CommandEnabled = true;
			CommandIcon = Resources.command_web;

		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
			doc.LoadHtml(v_HTMLVariable);

			var div = doc.DocumentNode.SelectSingleNode((string)await v_XPathQuery.EvaluateCode(engine));
			string divString = div.InnerText;

			divString.SetVariableValue(engine, v_OutputUserVariableName);
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_HTMLVariable", this, editor, 100, 300));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_XPathQuery", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultOutputGroupFor("v_OutputUserVariableName", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Query '{v_HTMLVariable}' - Store Result in '{v_OutputUserVariableName}']";
		}
	}
}