﻿using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Interfaces;
using OpenBots.Core.Properties;
using OpenBots.Core.Utilities.CommonUtilities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenBots.Commands.Image
{
    [Serializable]
	[Category("Image Commands")]
	[Description("This command captures an image on screen and stores it as a Bitmap variable.")]
	public class CaptureImageCommand : ScriptCommand, IImageCommands
	{

		[Required]
		[DisplayName("Capture Search Image")]
		[Description("Use the tool to capture an image that will be located on screen during execution.")]
		[SampleUsage("")]
		[Remarks("Images with larger color variance will be found more quickly than those with a lot of white space. \n" +
				 "For images that are primarily white space, tagging color to the top-left corner of the image and setting \n" +
				 "the relative click position will produce faster results.")]
		[Editor("ShowImageCaptureHelper", typeof(UIAdditionalHelperType))]
		public string v_ImageCapture { get; set; }

		[Required]
		[Editable(false)]
		[DisplayName("Output Image Variable")]
		[Description("Create a new variable or select a variable from the list.")]
		[SampleUsage("vUserVariable")]
		[Remarks("New variables/arguments may be instantiated by utilizing the Ctrl+K/Ctrl+J shortcuts.")]
		[CompatibleTypes(new Type[] { typeof(Bitmap) })]
		public string v_OutputUserVariableName { get; set; }

		public CaptureImageCommand()
		{
			CommandName = "CaptureImageCommand";
			SelectionName = "Capture Image";
			CommandEnabled = true;
			CommandIcon = Resources.command_camera;

		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;

			//user image to bitmap
			Bitmap capturedBmp = new Bitmap(CommonMethods.Base64ToImage(v_ImageCapture));
			capturedBmp.SetVariableValue(engine, v_OutputUserVariableName);
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			var imageCapture = commandControls.CreateDefaultPictureBoxFor("v_ImageCapture", this);
			

			RenderedControls.Add(commandControls.CreateDefaultLabelFor("v_ImageCapture", this));
			RenderedControls.AddRange(commandControls.CreateUIHelpersFor("v_ImageCapture", this, new Control[] { imageCapture }, editor));
			RenderedControls.Add(imageCapture);

			RenderedControls.AddRange(commandControls.CreateDefaultOutputGroupFor("v_OutputUserVariableName", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Capture Image on Screen - Store Image in '{v_OutputUserVariableName}']";
		}
	}
}
