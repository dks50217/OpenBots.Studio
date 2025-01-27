﻿//Copyright (c) 2019 Jason Bayldon
//Modifications - Copyright (c) 2020 OpenBots Inc.
//
//Licensed under the Apache License, Version 2.0 (the "License");
//you may not use this file except in compliance with the License.
//You may obtain a copy of the License at
//
//   http://www.apache.org/licenses/LICENSE-2.0
//
//Unless required by applicable law or agreed to in writing, software
//distributed under the License is distributed on an "AS IS" BASIS,
//WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//See the License for the specific language governing permissions and
//limitations under the License.
using OpenBots.Core.Enums;
using OpenBots.Core.Properties;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace OpenBots.Core.UI.Controls
{
    public partial class CommandItemControl : UserControl
    {      
        public UIAdditionalHelperType HelperType { get; set; }
        public object DataSource { get; set; }
        public string FunctionalDescription { get; set; }
        public string ImplementationDescription { get; set; }
        private string _commandDisplay;
        public string CommandDisplay
        {
            get
            {
                return _commandDisplay;
            }
            set
            {
                _commandDisplay = value;
                Invalidate();
            }
        }
        private Image commandImage;
        public Image CommandImage
        {
            get
            {
                return commandImage;
            }
            set
            {
                commandImage = value;
                Invalidate();
            }
        }

        public CommandItemControl()
        {
            InitializeComponent();
            Padding = new Padding(10, 0, 0, 0);
            ForeColor = Color.AliceBlue;
            Font = new Font("Segoe UI Semilight", 10);
            CommandImage = Resources.command_comment;
        }

        public CommandItemControl(string name, Bitmap commandImage, string commandDisplay)
        {
            InitializeComponent();
            Padding = new Padding(10, 0, 0, 0);
            ForeColor = Color.AliceBlue;
            Font = new Font("Segoe UI Semilight", 10);
            Name = name;
            CommandImage = commandImage;
            CommandDisplay = commandDisplay;
        }

        private void CommandItemControl_MouseEnter(object sender, EventArgs e)
        {
            Cursor = Cursors.Hand;
        }
        private void CommandItemControl_MouseLeave(object sender, EventArgs e)
        {
            Cursor = Cursors.Arrow;
        }

        private void CommandItemControl_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.DrawImage(CommandImage, 0, 0, 20, 20);
            e.Graphics.DrawString(CommandDisplay, Font, new SolidBrush(ForeColor), 18, 0);
        }

        private void CommandItemControl_Load(object sender, EventArgs e)
        {
        }
    }
}