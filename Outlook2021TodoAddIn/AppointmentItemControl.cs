using System;
using System.Drawing;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Outlook2021TodoAddIn
{
    public class AppointmentItemControl : Panel
    {
        private Outlook.AppointmentItem _appt;
        private Color _barColor;
        private bool _hasLocation;

        public AppointmentItemControl(Outlook.AppointmentItem appt, Font baseFont)
        {
            _appt = appt;
            _barColor = Color.SteelBlue;
            _hasLocation = !string.IsNullOrEmpty(appt.Location);

            if (!string.IsNullOrEmpty(appt.Categories))
            {
                try
                {
                    string firstCat = appt.Categories.Split(',')[0].Trim();
                    Outlook.Category cat = Globals.ThisAddIn.Application.Session.Categories[firstCat] as Outlook.Category;
                    if (cat != null)
                        _barColor = AppointmentsControl.TranslateCategoryColorStatic(cat.Color);
                }
                catch { }
            }

            this.Font = baseFont;
            this.BackColor = Color.White;
            this.Margin = new Padding(0, 0, 0, 3);

            // WICHTIG: kein Dock, feste Höhe
            int lineH = baseFont.Height + 2;
            this.Height = _hasLocation ? (lineH * 2 + 4) : (lineH + 4);
            this.Anchor = AnchorStyles.Left | AnchorStyles.Right;

            this.Paint += OnPaint;
            this.DoubleClick += (s, e) => { try { _appt.Display(false); } catch { } };
        }

        private void OnPaint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Color.White);

            int lineH = this.Font.Height + 2;
            int barX = 70;
            int barW = 4;
            int textX = barX + barW + 6;
            int textW = this.Width - textX - 2;

            // Kategorie-Tint
            if (_barColor != Color.SteelBlue)
            {
                Color tint = Color.FromArgb(60, _barColor.R, _barColor.G, _barColor.B);
                g.FillRectangle(new SolidBrush(tint), new Rectangle(textX - 2, 0, this.Width - textX + 2, this.Height));
            }

            // Uhrzeit
            if (!_appt.AllDayEvent)
            {
                using (var sf = new StringFormat { Alignment = StringAlignment.Far, LineAlignment = StringAlignment.Near })
                using (var brush = new SolidBrush(Color.Gray))
                    g.DrawString(_appt.Start.ToShortTimeString(), this.Font, brush, new RectangleF(0, 1, barX - 4, lineH), sf);
            }

            // Balken
            int barH = _hasLocation ? (lineH * 2 + 2) : (lineH + 2);
            g.FillRectangle(new SolidBrush(_barColor), new Rectangle(barX, 0, barW, barH));

            // Betreff (fett)
            using (var boldFont = new Font(this.Font, FontStyle.Bold))
            using (var sf = new StringFormat { Trimming = StringTrimming.EllipsisCharacter, FormatFlags = StringFormatFlags.NoWrap })
            using (var brush = new SolidBrush(Color.Black))
                g.DrawString(_appt.Subject, boldFont, brush, new RectangleF(textX, 1, textW, lineH), sf);

            // Ort
            if (_hasLocation)
            {
                using (var sf = new StringFormat { Trimming = StringTrimming.EllipsisCharacter, FormatFlags = StringFormatFlags.NoWrap })
                using (var brush = new SolidBrush(Color.Gray))
                    g.DrawString(_appt.Location, this.Font, brush, new RectangleF(textX, lineH + 2, textW, lineH), sf);
            }
        }
    }
}
