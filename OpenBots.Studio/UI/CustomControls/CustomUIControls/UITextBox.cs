﻿using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OpenBots.UI.CustomControls.CustomUIControls
{
    public partial class UITextBox : TextBox
    {
        const int WM_NCPAINT = 0x85;
        const int WM_PAINT = 0xf;
        const uint RDW_INVALIDATE = 0x1;
        const uint RDW_IUPDATENOW = 0x100;
        const uint RDW_FRAME = 0x400;

        [DllImport("user32.dll")]
        static extern IntPtr GetWindowDC(IntPtr hWnd);
        [DllImport("user32.dll")]
        static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);
        [DllImport("user32.dll")]
        static extern bool RedrawWindow(IntPtr hWnd, IntPtr lprc, IntPtr hrgn, uint flags);

        public bool IsEvaluateSnippet { get; set; }
        public bool ListBoxShown { get; set; }
        public int TabSize { get; set; } = 4;

        private Color _borderColor;
        public Color BorderColor
        {
            get 
            { 
                return _borderColor; 
            }
            set
            {
                _borderColor = value;
                RedrawWindow(Handle, IntPtr.Zero, IntPtr.Zero,
                    RDW_FRAME | RDW_IUPDATENOW | RDW_INVALIDATE);
            }
        }

        private bool _isDoubleBuffered;
        public bool IsDoubleBuffered
        {
            get 
            { 
                return _isDoubleBuffered; 
            }
            set
            {
                _isDoubleBuffered = value;
                if (_isDoubleBuffered == true)
                {
                    var dgvType = GetType();
                    var pi = dgvType.GetProperty("DoubleBuffered",
                          BindingFlags.Instance | BindingFlags.NonPublic);
                    pi.SetValue(this, true, null);
                }
            }
        }       

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if ((m.Msg == WM_NCPAINT || m.Msg == WM_PAINT) && BorderStyle == BorderStyle.Fixed3D)
            {
                var hdc = GetWindowDC(Handle);
                using (var g = Graphics.FromHdcInternal(hdc))
                using (var p = new Pen(BorderColor, 5))
                    g.DrawRectangle(p, new Rectangle(0, 0, Width - 1, Height - 1));
                ReleaseDC(Handle, hdc);
            }
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            RedrawWindow(Handle, IntPtr.Zero, IntPtr.Zero,
                   RDW_FRAME | RDW_IUPDATENOW | RDW_INVALIDATE);
        }

        protected override bool ProcessCmdKey(ref Message Msg, Keys keyData)
        {          
            if (ListBoxShown && keyData == Keys.Enter)
            {
                OnKeyDown(new KeyEventArgs(keyData));
                return true;
            }
            else if (keyData == Keys.Tab)
            {
                SelectedText += new string(' ', TabSize);
                return true;
            }
            return base.ProcessCmdKey(ref Msg, keyData);
        }

        protected override bool ProcessDialogKey(Keys keyData)
        {
            Keys key = (keyData & Keys.KeyCode);

            if (ListBoxShown && key == Keys.Enter)
            {
                OnKeyDown(new KeyEventArgs(key));
                return true;
            }
            else if (keyData == Keys.Tab)
            {
                SelectedText += new string(' ', TabSize);
                return true;
            }

            return base.ProcessDialogKey(keyData);
        }
    }
}
