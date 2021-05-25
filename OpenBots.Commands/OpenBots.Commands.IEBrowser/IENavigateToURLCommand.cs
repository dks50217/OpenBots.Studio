﻿using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Model.ApplicationModel;
using OpenBots.Core.Properties;
using OpenBots.Core.Utilities.CommonUtilities;

using SHDocVw;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenBots.Commands.IEBrowser
{
    [Serializable]
    [Category("IE Browser Commands")]
    [Description("This command navigates an existing IE web browser session to a given URL or web resource.")]
    public class IENavigateToURLCommand : ScriptCommand
    {
        [Required]
        [DisplayName("IE Browser Instance Name")]
        [Description("Enter the unique instance that was specified in the **IE Create Browser** command.")]
        [SampleUsage("MyIEBrowserInstance")]
        [Remarks("Failure to enter the correct instance name or failure to first call the **IE Create Browser** command will cause an error.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
        [CompatibleTypes(new Type[] { typeof(OBAppInstance) })]
        public string v_InstanceName { get; set; }

        [Required]
        [DisplayName("Navigate to URL")]
        [Description("Enter the destination URL that you want the IE instance to navigate to.")]
        [SampleUsage("\"https://example.com/\" || vURL")]
        [Remarks("")]
        [Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
        [CompatibleTypes(new Type[] { typeof(string) })]
        public string v_URL { get; set; }

        public IENavigateToURLCommand()
        {
            CommandName = "IENavigateToURLCommand";
            SelectionName = "IE Navigate to URL";           
            CommandEnabled = true;
            CommandIcon = Resources.command_web;

            v_InstanceName = "DefaultIEBrowser";
        }

        public async override Task RunCommand(object sender)
        {
            var engine = (IAutomationEngineInstance)sender;

            var browserObject = ((OBAppInstance)await v_InstanceName.EvaluateCode(engine)).Value;
            var browserInstance = (InternetExplorer)browserObject;

            browserInstance.Navigate((string)await v_URL.EvaluateCode(engine));
            IECreateBrowserCommand.WaitForReadyState(browserInstance);
        }

        public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
        {
            base.Render(editor, commandControls);

            RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
            RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_URL", this, editor));

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + $" [Navigate to '{v_URL}' - Instance Name '{v_InstanceName}']";
        }
    }
}