﻿using OpenBots.Commands.Server.Library;
using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Interfaces;
using OpenBots.Core.Properties;
using OpenBots.Core.Utilities.CommonUtilities;
using OpenBots.Server.SDK.HelperMethods;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenBots.Commands.QueueItem
{
	[Serializable]
	[Category("QueueItem Commands")]
	[Description("This command extends a QueueItem in an existing Queue in OpenBots Server.")]
	public class ExtendQueueItemCommand : ScriptCommand
	{
		[Required]
		[DisplayName("QueueItem")]
		[Description("Enter a QueueItem Dictionary variable.")]
		[SampleUsage("vQueueItem")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(Dictionary<string,object>) })]
		public string v_QueueItem { get; set; }

		public ExtendQueueItemCommand()
		{
			CommandName = "ExtendQueueItemCommand";
			SelectionName = "Extend QueueItem";
			CommandEnabled = true;
			CommandIcon = Resources.command_queueitem;

			CommonMethods.InitializeDefaultWebProtocol();
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var vQueueItem = (Dictionary<string, object>)await v_QueueItem.EvaluateCode(engine);

			Guid transactionKey = (Guid)vQueueItem["LockTransactionKey"];

			if (transactionKey == null || transactionKey == Guid.Empty)
				throw new NullReferenceException($"Transaction key {transactionKey} is invalid or not found");

			var userInfo = ServerSessionVariableMethods.GetUserInfo(engine);
			QueueItemMethods.ExtendQueueItem(userInfo, transactionKey);
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_QueueItem", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [QueueItem '{v_QueueItem}']";
		}
	}
}