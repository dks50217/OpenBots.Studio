﻿using OpenBots.Core.Command;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using Tasks = System.Threading.Tasks;

namespace OpenBots.Commands.Switch
{
    [Serializable]
    [Category("Switch Commands")]
    [Description("This command specifies the end of a switch block.")]
    public class EndSwitchCommand : ScriptCommand
    {
        public EndSwitchCommand()
        {
            CommandName = "EndSwitchCommand";
            SelectionName = "End Switch";
            CommandEnabled = true;
            CommandIcon = Resources.command_end_switch;

        }

        public async override Tasks.Task RunCommand(object sender)
        {
            //no execution required, used as a marker by the Automation Engine
        }

        public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
        {
            base.Render(editor, commandControls);

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue();
        }
    }
}
