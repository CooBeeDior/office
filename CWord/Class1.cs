using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Formatting;


namespace CWord
{
    public class Class1
    {
        /// <summary>
        /// https://www.e-iceblue.com/Tutorials/Spire.Doc/Spire.Doc-Program-Guide/Text/How-to-Insert-Text-to-Word-at-Exact-Position-in-C-VB.NET.html
        /// </summary>
        public void word()
        {
            Document doc = new Document();
            Section sec = doc.AddSection();
            Paragraph par = sec.AddParagraph();

            TextBox textBox = par.AppendTextBox(180, 30);
            textBox.Format.VerticalOrigin = VerticalOrigin.Margin;
            textBox.Format.VerticalPosition = 100;
            textBox.Format.HorizontalOrigin = HorizontalOrigin.Margin;
            textBox.Format.HorizontalPosition = 50;
            textBox.Format.NoLine = true;
            CharacterFormat format = new CharacterFormat(doc);

            format.FontName = "Calibri";
            format.FontSize = 15;
            format.Bold = true; 
            Paragraph par1 = textBox.Body.AddParagraph();
            par1.AppendText("This is my new string").ApplyCharacterFormat(format); 
            doc.SaveToFile("result.docx", FileFormat.Docx);

        }
    }
}
