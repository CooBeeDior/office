using Spire.Pdf;
using Spire.Pdf.Graphics;
using System;
using System.Drawing;

namespace CPdf
{
    public class Class1
    {
        /// <summary>
        /// https://www.e-iceblue.com/Tutorials/Spire.PDF/Spire.PDF-Program-Guide/Convert-HTML-to-PDF-Customize-HTML-to-PDF-Conversion-by-Yourself.html
        /// </summary>
        public void pdf()
        { 
            PdfDocument doc = new PdfDocument();  
            PdfPageBase page = doc.Pages.Add();   
            PdfGraphicsState state = page.Canvas.Save();  
            PdfFont font = new PdfFont(PdfFontFamily.Helvetica, 10f); 
            PdfSolidBrush brush = new PdfSolidBrush(Color.Blue); 
            PdfStringFormat centerAlignment        = new PdfStringFormat(PdfTextAlignment.Left, PdfVerticalAlignment.Middle);

            float x = page.Canvas.ClientSize.Width / 2; 
            float y = 380; 
            page.Canvas.TranslateTransform(x, y); 
            for (int i = 0; i < 12; i++) 
            { 
                page.Canvas.RotateTransform(30); 
                page.Canvas.DrawString("Go! Turn Around! Go! Go! Go!", font, brush, 20, 0, centerAlignment); 
            } 
            page.Canvas.Restore(state); 
            doc.SaveToFile("DrawText.pdf"); 
            doc.Close(); 
        }
    }
}
