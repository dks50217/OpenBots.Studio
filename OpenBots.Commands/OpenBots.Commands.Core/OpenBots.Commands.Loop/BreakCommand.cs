﻿using OpenBots.Core.Command;
using OpenBots.Core.Interfaces;
using OpenBots.Core.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;

namespace OpenBots.Commands.Loop
{
    [Serializable]
    [Category("Loop Commands")]
    [Description("This command signifies that the current loop should exit and resume execution outside the current loop.")]
    public class BreakCommand : ScriptCommand
    {
        public BreakCommand()
        {
            CommandName = "BreakCommand";
            SelectionName = "Break";
            CommandEnabled = true;
            CommandIcon = Resources.command_exitloop;
        }

        public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
        {
            base.Render(editor, commandControls);

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return "Break";
        }
    }
}