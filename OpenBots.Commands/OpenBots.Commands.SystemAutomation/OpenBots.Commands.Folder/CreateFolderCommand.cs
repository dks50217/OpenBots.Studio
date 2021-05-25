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
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenBots.Commands.Folder
{
	[Serializable]
	[Category("Folder Operation Commands")]
	[Description("This command creates a folder in a specified location.")]
	public class CreateFolderCommand : ScriptCommand
	{
		[Required]
		[DisplayName("New Folder Name")]
		[Description("Enter the name of the new folder.")]
		[SampleUsage("\"myFolderName\" || vFolderName")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_NewFolderName { get; set; }

		[Required]
		[DisplayName("Directory Path")]
		[Description("Enter or Select the path to the directory to create the folder in.")]
		[SampleUsage("@\"C:\\temp\" || ProjectPath + @\"\\temp\" || vDirectoryPath")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[Editor("ShowFolderSelectionHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_DestinationDirectory { get; set; }

		[Required]
		[DisplayName("Delete Existing Folder")]
		[PropertyUISelectionOption("Yes")]
		[PropertyUISelectionOption("No")]
		[Description("Specify whether the folder should be deleted first if it already exists.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_DeleteExisting { get; set; }

		public CreateFolderCommand()
		{
			CommandName = "CreateFolderCommand";
			SelectionName = "Create Folder";
			CommandEnabled = true;
			CommandIcon = Resources.command_folders;
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			//apply variable logic
			var destinationDirectory = (string)await v_DestinationDirectory.EvaluateCode(engine);
			var newFolder = (string)await v_NewFolderName.EvaluateCode(engine);

            if (!Directory.Exists(destinationDirectory))
            {
				throw new DirectoryNotFoundException($"Directory {destinationDirectory} is not a valid directory");
            }

			var finalPath = Path.Combine(destinationDirectory, newFolder);
			//delete folder if it exists AND the delete option is selected 
			if (v_DeleteExisting == "Yes" && Directory.Exists(finalPath))
			{
				Directory.Delete(finalPath, true);
			}

			//create folder if it doesn't exist
			if (!Directory.Exists(finalPath))
			{
				Directory.CreateDirectory(finalPath);
			}
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_NewFolderName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_DestinationDirectory", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_DeleteExisting", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $"[Folder Path '{v_DestinationDirectory}\\{v_NewFolderName}']";
		}
	}
}