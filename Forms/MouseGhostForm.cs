using System;
using System.Drawing;
using System.Windows.Forms;

namespace SW2026RibbonAddin.Forms
{
    /// <summary>
    /// A lightweight, click-through overlay that follows the mouse and draws
    /// a rectangle indicating where the note will be placed.
    /// </summary>
    internal sealed class MouseGhostForm : Form
    {
        private readonly Timer _timer;
        private Size _rectSize = new Size(160, 60); // default
        private const int OFFSET = 18;

        public MouseGhostForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            BackColor = Color.Magenta;           // will be transparent
            TransparencyKey = Color.Magenta;
            Opacity = 0.85;                      // slight translucency
            DoubleBuffered = true;
            Enabled = false;                     // don't steal focus/clicks

            _timer = new Timer();
            _timer.Interval = 16; // ~60 FPS
            _timer.Tick += (s, e) =>
            {
                var pos = Cursor.Position;
                Location = new Point(pos.X + OFFSET, pos.Y + OFFSET);
            };

            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.UserPaint, true);
        }

        /// <summary>
        /// Measure the text with given font to size the rectangle similar to SW notes.
        /// </summary>
        public void InitializeForText(string text, string fontName, float fontSizePts)
        {
            if (fontSizePts <= 0) fontSizePts = 12f;
            if (string.IsNullOrWhiteSpace(fontName)) fontName = "Tahoma";

            using (var bmp = new Bitmap(1, 1))
            using (var g = Graphics.FromImage(bmp))
            using (var font = new Font(fontName, fontSizePts, GraphicsUnit.Point))
            {
                var sz = TextRenderer.MeasureText(g, string.IsNullOrEmpty(text) ? " " : text, font,
                                                  new Size(int.MaxValue, int.MaxValue),
                                                  TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding);
                // pad a little to mimic SW note bounding box
                int w = Math.Max(120, sz.Width + 12);
                int h = Math.Max((int)(fontSizePts * 2.0f), sz.Height + 8);
                _rectSize = new Size(w, h);
                Size = _rectSize;
            }
        }

        public void StartFollowing()
        {
            // set initial size/pos
            Size = _rectSize;
            var pos = Cursor.Position;
            Location = new Point(pos.X + OFFSET, pos.Y + OFFSET);
            Show();
            _timer.Start();
        }

        protected override CreateParams CreateParams
        {
            get
            {
                const int WS_EX_TOOLWINDOW = 0x00000080;
                const int WS_EX_TRANSPARENT = 0x00000020;
                const int WS_EX_LAYERED = 0x00080000;

                var cp = base.CreateParams;
                cp.ExStyle |= WS_EX_TOOLWINDOW | WS_EX_TRANSPARENT | WS_EX_LAYERED;
                return cp;
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            var rect = new Rectangle(0, 0, _rectSize.Width - 1, _rectSize.Height - 1);

            // Fill with a faint hatch-like alpha
            using (var brush = new SolidBrush(Color.FromArgb(24, 0, 0, 0)))
                g.FillRectangle(brush, rect);

            // Outline
            using (var pen = new Pen(Color.Black, 1))
                g.DrawRectangle(pen, rect);
            using (var pen = new Pen(Color.White, 1))
            {
                var inner = Rectangle.Inflate(rect, -1, -1);
                g.DrawRectangle(pen, inner);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _timer?.Stop();
                _timer?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
