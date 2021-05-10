﻿using Microsoft.Office.Interop.Word;
using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Model.ApplicationModel;
using OpenBots.Core.Properties;
using OpenBots.Core.Utilities.CommonUtilities;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Word.Application;
using Tasks = System.Threading.Tasks;

namespace OpenBots.Commands.Word
{
	[Serializable]
	[Category("Word Commands")]
	[Description("This command appends an image to a Word Document.")]
	public class WordAppendImageCommand : ScriptCommand
	{
		[Required]
		[DisplayName("Word Instance Name")]
		[Description("Enter the unique instance that was specified in the **Create Application** command.")]
		[SampleUsage("MyWordInstance")]
		[Remarks("Failure to enter the correct instance or failure to first call the **Create Application** command will cause an error.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(OBAppInstance) })]
		public string v_InstanceName { get; set; }

		[Required]
		[DisplayName("Image File Path")]    
		[Description("Enter the file path of the image to append to the Document.")]
		[SampleUsage("@\"C:\\temp\\myfile.png\" || ProjectPath + @\"\\myfile.png\" || vFilePath")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[Editor("ShowFileSelectionHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_ImagePath { get; set; }

		public WordAppendImageCommand()
		{
			CommandName = "WordAppendImageCommand";
			SelectionName = "Append Image";
			CommandEnabled = true;
			CommandIcon = Resources.command_files;

			v_InstanceName = "DefaultWord";
		}

		public async override Tasks.Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var vImagePath = (string)await v_ImagePath.EvaluateCode(engine);
			var wordObject = ((OBAppInstance)await v_InstanceName.EvaluateCode(engine)).Value;

			Application wordInstance = (Application)wordObject;
			Document wordDocument = wordInstance.ActiveDocument;

			//Appends image after text/images
			object collapseEnd = WdCollapseDirection.wdCollapseEnd;
			Range imageRange = wordDocument.Content;
			imageRange.Collapse(ref collapseEnd);
			imageRange.InlineShapes.AddPicture(vImagePath, Type.Missing, Type.Missing, imageRange);

			Paragraph paragraph = wordDocument.Content.Paragraphs.Add();
			paragraph.Format.SpaceAfter = 10f;
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_ImagePath", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Append '{v_ImagePath}' - Instance Name '{v_InstanceName}']";
		}
	}
}