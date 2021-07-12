﻿using Open3270;
using OpenBots.Commands.Terminal.Library;
using OpenBots.Core.User32;
using System;
using System.Drawing;
using System.Security;
using System.Text;
using System.Windows.Forms;

namespace OpenBots.Commands.Terminal.Forms
{
    public class OpenEmulator : RichTextBox, IAudit
    {
        public TNEmulator TN3270 = new TNEmulator();      
        private bool _isRedrawing = false;
        public Point Coordinates { get; set; } = new Point(0, 0);
        public int FieldIndex { get; set; } = 0;
        private Color _textColor = Color.Lime;
        public string Username { get; set; }
        public SecureString Password { get; set; }

        public delegate bool ConnectDelegate(string server, int port, string type, bool useSsl);
        public bool Connect(string server, int port, string type, bool useSsl)
        {
            if (InvokeRequired)
            {
                var d = new ConnectDelegate(Connect);
                return (bool)Invoke(d, new object[] { server, port, type, useSsl });
            }
            else
            {
                try
                {
                    TN3270.Config.UseSSL = useSsl;
                    TN3270.Config.TermType = type;
                    TN3270.Audit = this;
                    //TN3270.Debug = true;
                    TN3270.Config.FastScreenMode = true;

                    TN3270.Connect(server, port, string.Empty);

                    Redraw();                   
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error");
                    return false;
                }
                return true;
            }                 
        }

        public delegate void DisconnectDelegate();
        public void Disconnect()
        {
            if (InvokeRequired)
            {
                var d = new DisconnectDelegate(Disconnect);
                Invoke(d, new object[] { });
            }
            else
            {
                TN3270.Close();
                Rtf = "";
            }          
        }


        public delegate void RedrawDelegate();
        public void Redraw()
        {
            if (InvokeRequired)
            {
                var d = new RedrawDelegate(Redraw);
                Invoke(d, new object[] { });
            }
            else
            {
                _isRedrawing = true;

                TN3270.WaitTillKeyboardUnlocked(2000);
                Text = TN3270.CurrentScreenXML.Dump();

                SelectAll();
                SelectionColor = Color.Lime;

                if (TN3270.CurrentScreenXML.Fields == null)
                {
                    SelectionProtected = false;
                    SelectionColor = Color.Lime;
                    DeselectAll();

                    for (int i = 0; i < Text.Length; i++)
                    {
                        Select(i, 1);
                        if (SelectedText != " " && SelectedText != "\n")
                            SelectionColor = Color.Lime;
                    }
                    return;
                }

                SelectionProtected = true;
                foreach (Open3270.TN3270.XMLScreenField field in TN3270.CurrentScreenXML.Fields)
                {
                    if (string.IsNullOrEmpty(field.Text))
                        continue;

                    Application.DoEvents();
                    _textColor = Color.Lime;
                    if (field.Attributes.FieldType == "High" && field.Attributes.Protected)
                        _textColor = Color.White;
                    else if (field.Attributes.FieldType == "High")
                        _textColor = Color.Red;
                    else if (field.Attributes.Protected)
                        _textColor = Color.RoyalBlue;

                    Select((field.Location.top * 84 + 172) + field.Location.left + 3, field.Location.length);
                    SelectionProtected = false;
                    SelectionColor = _textColor;

                    if (_textColor == Color.White || _textColor == Color.RoyalBlue)
                        SelectionProtected = true;
                }
                _isRedrawing = false;

                Select((TN3270.CursorY * 84 + 172) + TN3270.CursorX + 3, 0);
            }
            
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (TN3270.IsConnected)
            {
                if (TerminalKeys.Open3270TerminalKeysDict.ContainsKey(e.KeyCode))
                {
                    TnKey sendKey = TerminalKeys.Open3270TerminalKeysDict[e.KeyCode];
                    TN3270.SendKey(false, sendKey, 2000);
                }
                else if (e.KeyCode == Keys.ShiftKey)
                    return;
                else
                {
                    bool shift = false;

                    var character = (char)e.KeyCode;

                    if (GlobalHook.IsKeyDown(Keys.ShiftKey))
                    {
                        if (GlobalHook.IsKeyToggled(Keys.Capital))
                        {
                            if (char.IsLetter(character))
                                shift = false;
                            else
                                shift = true;
                        }
                        else
                            shift = true;
                    }
                    else
                    {
                        if (GlobalHook.IsKeyToggled(Keys.Capital))
                        {
                            if (char.IsLetter(character))
                                shift = true;
                        }
                        else
                            shift = false;
                    }

                    var buf = new StringBuilder(256);
                    var keyboardState = new byte[256];

                    if (shift)
                        keyboardState[(int)Keys.ShiftKey] = 0xff;

                    GlobalHook.ToUnicode((uint)e.KeyCode, 0, keyboardState, buf, 256, 0);
                    var selectedKey = buf.ToString();
                    TN3270.SetText(selectedKey);
                }
                Redraw();
            }
           e.Handled = true;
                
        }

        protected override void OnKeyPress(KeyPressEventArgs e)
        {
            e.Handled = true;
            base.OnKeyPress(e);
        }

        protected override void OnSelectionChanged(EventArgs e)
        {
            if (TN3270.IsConnected)
            {
                if (!_isRedrawing)
                {
                    int i = SelectionStart - 172 , x, y = 0;
                    while (i >= 84)
                    {
                        y++;
                        i -= 84;
                    }
                    x = i - 3;

                    if (x < 0 || SelectionStart <= 174 || SelectionStart > 2271)
                    {
                        x = -1;
                        y = -1;
                    }
                        
                    Coordinates = new Point(x, y);
                    TN3270.SetCursor(Coordinates.X, Coordinates.Y);

                    FieldIndex = -1;

                    foreach (Open3270.TN3270.XMLScreenField field in TN3270.CurrentScreenXML.Fields)
                    { 
                        if ((Coordinates.Y * 80 + Coordinates.X) >= field.Location.position &&
                            (Coordinates.Y * 80 + Coordinates.X) < field.Location.position + field.Location.length)
                        {
                            FieldIndex = Array.IndexOf(TN3270.CurrentScreenXML.Fields, field);                           
                        }
                    }
                    base.OnSelectionChanged(e);
                }
            }
        }

        public void InitializeComponent()
        {
            SuspendLayout();

            Font = new Font("Consolas", 10F, FontStyle.Regular);
            Size = new Size(Font.Height * 39, Font.Height * 31);
            ResumeLayout(false);
        }

        public void WriteLine(string text)
        {
            Console.WriteLine(text);
        }

        public void Write(string text)
        {
            Console.Write(text);
        }

        public delegate void DisposeTerminalDelegate();
        public void DisposeTerminal()
        {
            if (InvokeRequired)
            {
                var d = new DisposeTerminalDelegate(DisposeTerminal);
                Invoke(d, new object[] { });
            }
            else
            {
                Dispose();
            }
        }

        public void PromptForUsername()
        {
            frmTerminalInputBox inputBox = new frmTerminalInputBox("Enter the username:", "Username");
            if (inputBox.ShowDialog() == DialogResult.OK)
            {
                TN3270.SetText(inputBox.txtInput.Text);
                int result = TN3270.WaitForTextOnScreen(5000, inputBox.txtInput.Text);

                if (result == -1)
                    throw new TimeoutException($"Unable to find '{inputBox.txtInput.Text}' within the allotted time of 5 seconds.");

                Redraw();
            }
        }

        public void PromptForPassword()
        {
            frmTerminalInputBox inputBox = new frmTerminalInputBox("Enter the password:", "Password", true);
            if (inputBox.ShowDialog() == DialogResult.OK)
            {
                TN3270.SetText(inputBox.txtInput.Text);
                int result = TN3270.WaitForTextOnScreen(5000, inputBox.txtInput.Text);

                if (result == -1)
                    throw new TimeoutException($"Unable to find the password text within the allotted time of 5 seconds.");

                Redraw();
            }
        }
    }
}
