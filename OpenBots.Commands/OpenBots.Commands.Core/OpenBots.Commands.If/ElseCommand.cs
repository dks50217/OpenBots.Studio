﻿using OpenBots.Core.Command;
using OpenBots.Core.Interfaces;
using OpenBots.Core.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;

namespace OpenBots.Commands.If
{
    [Serializable]
    [Category("If Commands")]
    [Description("This command declares the seperation between the actions based on the 'true' or 'false' condition.")]
    public class ElseCommand : ScriptCommand
    {
        public ElseCommand()
        {
            CommandName = "ElseCommand";
            SelectionName = "Else";
            CommandEnabled = true;
            CommandIcon = Resources.command_else;

        }

        public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
        {
            base.Render(editor, commandControls);

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return "Else";
        }
    }
}