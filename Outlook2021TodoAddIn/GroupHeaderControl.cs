using System;
using System.Drawing;
using System.Windows.Forms;

namespace Outlook2021TodoAddIn
{
    public class GroupHeaderControl : Panel
    {
        private string _text;

        public GroupHeaderControl(string text, Font baseFont)
        {
            _text = text;
            this.Font = baseFont;
            this.BackColor = Color.White;
            this.Margin = new Padding(0, 4, 0, 0);
            this.Height = baseFont.Height + 6;
            // WICHTIG: kein Dock
            this.Paint += OnPaint;
        }

        private void OnPaint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Color.White);

            using (var boldFont = new Font(this.Font, FontStyle.Bold))
            using (var brush = new SolidBrush(Color.SteelBlue))
            {
                SizeF textSize = g.MeasureString(_text, boldFont);
                g.DrawString(_text, boldFont, brush, new PointF(0, 2));
                int lineX = (int)textSize.Width + 4;
                int lineY = this.Height / 2;
                using (var pen = new Pen(Color.LightSteelBlue, 1))
                    g.DrawLine(pen, lineX, lineY, this.Width - 2, lineY);
            }
        }
    }
}
