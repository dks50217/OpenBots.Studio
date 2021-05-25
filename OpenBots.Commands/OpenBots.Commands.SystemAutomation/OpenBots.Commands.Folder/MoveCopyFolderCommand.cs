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
	[Description("This command moves/copies a folder to a specified location.")]
	public class MoveCopyFolderCommand : ScriptCommand
	{
		[Required]
		[DisplayName("Folder Operation Type")]
		[PropertyUISelectionOption("Move Folder")]
		[PropertyUISelectionOption("Copy Folder")]
		[Description("Specify whether you intend to move or copy the folder.")]
		[SampleUsage("")]
		[Remarks("Moving will remove the folder from the original path while Copying will not.")]
		public string v_OperationType { get; set; }

		[Required]
		[DisplayName("Source Folder Path")]
		[Description("Enter or Select the path to the original folder.")]
		[SampleUsage("@\"C:\\temp\" || ProjectPath + @\"\\temp\" || vDirectoryPath")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[Editor("ShowFolderSelectionHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_SourceFolderPath { get; set; }

		[Required]
		[DisplayName("Destination Folder Path")]
		[Description("Enter or Select the destination folder path.")]
		[SampleUsage("@\"C:\\temp\" || ProjectPath + @\"\\temp\" || vDirectoryPath")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[Editor("ShowFolderSelectionHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_DestinationDirectory { get; set; }

		[Required]
		[DisplayName("Create Destination Folder")]
		[PropertyUISelectionOption("Yes")]
		[PropertyUISelectionOption("No")]
		[Description("Specify whether the destination directory should be created if it does not already exist.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_CreateDirectory { get; set; }

		[Required]
		[DisplayName("Delete Existing Folder")]
		[PropertyUISelectionOption("Yes")]
		[PropertyUISelectionOption("No")]
		[Description("Specify whether the folder should be deleted first if it already exists in the destination directory.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_DeleteExisting { get; set; }

		public MoveCopyFolderCommand()
		{
			CommandName = "MoveCopyFolderCommand";
			SelectionName = "Move/Copy Folder";
			CommandEnabled = true;
			CommandIcon = Resources.command_folders;

			v_CreateDirectory = "Yes";
			v_DeleteExisting = "Yes";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			//apply variable logic
			var sourceFolder = (string)await v_SourceFolderPath.EvaluateCode(engine);
			var destinationFolder = (string)await v_DestinationDirectory.EvaluateCode(engine);
			
			if (!Directory.Exists(sourceFolder))
            {
				throw new DirectoryNotFoundException($"Directory {sourceFolder} does not exist");
            }

			if ((v_CreateDirectory == "Yes") && (!Directory.Exists(destinationFolder)))
			{
				Directory.CreateDirectory(destinationFolder);
			}

			//get source folder name and info
			DirectoryInfo sourceFolderInfo = new DirectoryInfo(sourceFolder);

			//create final path
			var finalPath = Path.Combine(destinationFolder, sourceFolderInfo.Name);

			//delete if it already exists per user
			if (v_DeleteExisting == "Yes" && Directory.Exists(finalPath))
			{
				Directory.Delete(finalPath, true);
			}

			if (v_OperationType == "Move Folder")
			{
				//move folder
				Directory.Move(sourceFolder, finalPath);
			}
			else if (v_OperationType == "Copy Folder")
			{
				//copy folder
				DirectoryCopy(sourceFolder, finalPath, true);   
			}
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_OperationType", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_SourceFolderPath", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_DestinationDirectory", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_CreateDirectory", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_DeleteExisting", this, editor));
			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [{v_OperationType} From '{v_SourceFolderPath}' to '{v_DestinationDirectory}']";
		}

		private void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
		{
			// Get the subdirectories for the specified directory.
			DirectoryInfo dir = new DirectoryInfo(sourceDirName);

			if (!dir.Exists)
			{
				throw new DirectoryNotFoundException(
					"Source directory does not exist or could not be found: "
					+ sourceDirName);
			}

			DirectoryInfo[] dirs = dir.GetDirectories();
			// If the destination directory doesn't exist, create it.
			Directory.GetParent(destDirName);
			if (!Directory.GetParent(destDirName).Exists)
			{
				throw new DirectoryNotFoundException(
					"Destination directory does not exist or could not be found: "
					+ Directory.GetParent(destDirName));
			}

			if (!Directory.Exists(destDirName))
			{
				Directory.CreateDirectory(destDirName);
			}

			// Get the files in the directory and copy them to the new location.
			FileInfo[] files = dir.GetFiles();
			foreach (FileInfo file in files)
			{
				string temppath = Path.Combine(destDirName, file.Name);
				file.CopyTo(temppath, false);
			}

			// If copying subdirectories, copy them and their contents to new location.
			if (copySubDirs)
			{
				foreach (DirectoryInfo subdir in dirs)
				{
					string temppath = Path.Combine(destDirName, subdir.Name);
					DirectoryCopy(subdir.FullName, temppath, copySubDirs);
				}
			}
		}
	}
}