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
using System.Windows.Forms;
using Tasks = System.Threading.Tasks;

namespace OpenBots.Commands.ErrorHandling
{
	[Serializable]
	[Category("Error Handling Commands")]
	[Description("This command throws an exception during script execution.")]
	public class ThrowCommand : ScriptCommand
	{
		[Required]
		[DisplayName("Exception Type")]
		[PropertyUISelectionOption("AccessViolationException")]
		[PropertyUISelectionOption("ArgumentException")]
		[PropertyUISelectionOption("ArgumentNullException")]
		[PropertyUISelectionOption("ArgumentOutOfRangeException")]
		[PropertyUISelectionOption("DivideByZeroException")]
		[PropertyUISelectionOption("Exception")]
		[PropertyUISelectionOption("FileNotFoundException")]
		[PropertyUISelectionOption("FormatException")]
		[PropertyUISelectionOption("IndexOutOfRangeException")]
		[PropertyUISelectionOption("InvalidDataException")]
		[PropertyUISelectionOption("InvalidOperationException")]
		[PropertyUISelectionOption("KeyNotFoundException")]
		[PropertyUISelectionOption("NotSupportedException")]
		[PropertyUISelectionOption("NullReferenceException")]
		[PropertyUISelectionOption("OverflowException")]
		[PropertyUISelectionOption("TimeoutException")]
		[Description("Select the appropriate exception type to throw.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_ExceptionType { get; set; }

		[Required]
		[DisplayName("Exception Message")]
		[Description("Enter a custom exception message.")]
		[SampleUsage("\"A Custom Message\" || vExceptionMessage")]
		[Remarks("The selected exception with this custom message will be thrown.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_ExceptionMessage { get; set; }

		public ThrowCommand()
		{
			CommandName = "ThrowCommand";
			SelectionName = "Throw";
			CommandEnabled = true;
			CommandIcon = Resources.command_exception;

			v_ExceptionType = "Exception";
		}

		public async override Tasks.Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var exceptionMessage = (string)await v_ExceptionMessage.EvaluateCode(engine);

			Exception ex;
			switch(v_ExceptionType)
			{
				case "AccessViolationException":
					ex = new AccessViolationException(exceptionMessage);
					break;
				case "ArgumentException":
					ex = new ArgumentException(exceptionMessage);
					break;
				case "ArgumentNullException":
					ex = new ArgumentNullException(exceptionMessage);
					break;
				case "ArgumentOutOfRangeException":
					ex = new ArgumentOutOfRangeException(exceptionMessage);
					break;
				case "DivideByZeroException":
					ex = new DivideByZeroException(exceptionMessage);
					break;
				case "Exception":
					ex = new Exception(exceptionMessage);
					break;
				case "FileNotFoundException":
					ex = new FileNotFoundException(exceptionMessage);
					break;
				case "FormatException":
					ex = new FormatException(exceptionMessage);
					break;
				case "IndexOutOfRangeException":
					ex = new IndexOutOfRangeException(exceptionMessage);
					break;
				case "InvalidDataException":
					ex = new InvalidDataException(exceptionMessage);
					break;
				case "InvalidOperationException":
					ex = new InvalidOperationException(exceptionMessage);
					break;
				case "KeyNotFoundException":
					ex = new KeyNotFoundException(exceptionMessage);
					break;
				case "NotSupportedException":
					ex = new NotSupportedException(exceptionMessage);
					break;
				case "NullReferenceException":
					ex = new NullReferenceException(exceptionMessage);
					break;
				case "OverflowException":
					ex = new OverflowException(exceptionMessage);
					break;
				case "TimeoutException":
					ex = new TimeoutException(exceptionMessage);
					break;
				default:
					throw new NotImplementedException("Selected exception type " + v_ExceptionType + " is not implemented.");
			}
			throw ex;
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_ExceptionType", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_ExceptionMessage", this, editor));
			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Exception '{v_ExceptionType}' - Message '{v_ExceptionMessage}']";
		}
	}
}