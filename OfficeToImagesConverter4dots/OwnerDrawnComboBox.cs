using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.ComponentModel;
using System.Drawing;

namespace OfficeToImagesConverter4dots
{
    [ToolboxBitmap(typeof(ComboBox))]
    public class OwnerDrawnComboBox : ComboBox
    {
        public OwnerDrawnComboBox() : base()
        {
            this.DrawMode = DrawMode.OwnerDrawFixed;
            //this.SetStyle(ControlStyles.AllPaintingInWmPaint, true);

            this.DropDownBorderColor = Color.Black;
            this.DropDownBackColor = SystemColors.Control;
        }

        [Category("Appearance")]
        [Description("The border color of the drop down list")]
        [DefaultValue(typeof(Color), "Red")]
        public Color DropDownBorderColor { get; set; }

        [Category("Appearance")]
        [Description("The background color of the drop down list")]
        [DefaultValue(typeof(Color), "Yellow")]
        public Color DropDownBackColor { get; set; }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            return;
            
            string txt = this.Text;
            
            int epos = txt.IndexOf(" - ") + 1;

            if (txt.IndexOf(" - ") < 0)
            {
                SizeF sz = e.Graphics.MeasureString(txt, this.Font);
                e.Graphics.DrawString(txt, this.Font, Brushes.Black,
                  new RectangleF(e.ClipRectangle.X, e.ClipRectangle.Y, sz.Width, e.ClipRectangle.Height));
            }
            else
            {
                string first = txt.Substring(0, epos);
                string second = txt.Substring(epos);

                Font fsec = new Font(this.Font, FontStyle.Regular);

                SizeF sz = e.Graphics.MeasureString(first, this.Font);
                e.Graphics.DrawString(first, this.Font, Brushes.Black,
                  new RectangleF(e.ClipRectangle.X, e.ClipRectangle.Y, sz.Width, e.ClipRectangle.Height));

                SizeF sz2 = e.Graphics.MeasureString(second, fsec);

                e.Graphics.DrawString(second, fsec, Brushes.Black,
                  new RectangleF(e.ClipRectangle.X + sz.Width, e.ClipRectangle.Y, sz.Width + sz2.Width, e.ClipRectangle.Height));
            }
        }

        protected override void OnDrawItem(DrawItemEventArgs e)
        {                        
            if (e.Index < 0)
                return;

            if (e.Index != 0)
            {
                int a = 0;
            }
            base.OnDrawItem(e);

            //e.DrawBackground();

            if ((e.State & DrawItemState.ComboBoxEdit) == DrawItemState.ComboBoxEdit)
            {
               //return;
            }            
            // Draw the background of the item.
            else if (
                ((e.State & DrawItemState.Focus) == DrawItemState.Focus) ||
                ((e.State & DrawItemState.Selected) == DrawItemState.Selected) ||
                ((e.State & DrawItemState.HotLight) == DrawItemState.HotLight)
               )
            {
                e.DrawBackground();
            }
            else
            {
               
                Brush backgroundBrush = new SolidBrush(DropDownBackColor);
                e.Graphics.FillRectangle(backgroundBrush, e.Bounds);
                 
                 
            }

            //Draw item text
            string txt = Items[e.Index].ToString();
            int epos = txt.IndexOf(" - ") + 1;

            Font fsec = new Font(this.Font, FontStyle.Regular);

            if (txt.IndexOf(" - ") < 0)
            {
                SizeF sz = e.Graphics.MeasureString(txt, fsec);
                e.Graphics.DrawString(txt, fsec, Brushes.Black,
                  new RectangleF(e.Bounds.X, e.Bounds.Y, sz.Width, e.Bounds.Height));
                /*
                SizeF sz = e.Graphics.MeasureString(txt, this.Font);
                e.Graphics.DrawString(txt, this.Font, Brushes.Black,
                  new RectangleF(e.Bounds.X, e.Bounds.Y, sz.Width, e.Bounds.Height));
                 */
            }
            else
            {
                string first = txt.Substring(0, epos);
                string second = txt.Substring(epos);

                

                SizeF sz = e.Graphics.MeasureString(first, this.Font);

                
                e.Graphics.DrawString(first, this.Font, Brushes.Black,
                      new RectangleF(e.Bounds.X, e.Bounds.Y, sz.Width, e.Bounds.Height));
                
                SizeF sz2 = e.Graphics.MeasureString(second, fsec);

                e.Graphics.DrawString(second, fsec, Brushes.Black,
                  new RectangleF(e.Bounds.X + sz.Width, e.Bounds.Y, sz.Width + sz2.Width, e.Bounds.Height));
            }

            // Draw the focus rectangle if the mouse hovers over an item.
            if ((e.State & DrawItemState.Focus) == DrawItemState.Focus)
                e.DrawFocusRectangle();

            if ((e.State & DrawItemState.ComboBoxEdit) != DrawItemState.ComboBoxEdit)            
            {
                //Draw the border around the whole DropDown area
                Pen borderPen = new Pen(DropDownBorderColor, 1);
                Point start;
                Point end;

                if (e.Index == 0)
                {
                    //Draw top border
                    start = new Point(e.Bounds.Left, e.Bounds.Top);
                    end = new Point(e.Bounds.Left + e.Bounds.Width - 1, e.Bounds.Top);
                    e.Graphics.DrawLine(borderPen, start, end);
                }

                //Draw left border
                start = new Point(e.Bounds.Left, e.Bounds.Top);
                end = new Point(e.Bounds.Left, e.Bounds.Top + e.Bounds.Height - 1);
                e.Graphics.DrawLine(borderPen, start, end);

                //Draw Right border
                start = new Point(e.Bounds.Left + e.Bounds.Width - 1, e.Bounds.Top);
                end = new Point(e.Bounds.Left + e.Bounds.Width - 1, e.Bounds.Top + e.Bounds.Height - 1);
                e.Graphics.DrawLine(borderPen, start, end);

                if (e.Index == Items.Count - 1)
                {
                    //Draw bottom border
                    start = new Point(e.Bounds.Left, e.Bounds.Top + e.Bounds.Height - 1);
                    end = new Point(e.Bounds.Left + e.Bounds.Width - 1, e.Bounds.Top + e.Bounds.Height - 1);
                    e.Graphics.DrawLine(borderPen, start, end);
                }
            }
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            e.SuppressKeyPress = true;

            base.OnKeyDown(e);
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // OwnerDrawnComboBox
            // 
            this.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;

            this.ResumeLayout(false);

        }
    }
}