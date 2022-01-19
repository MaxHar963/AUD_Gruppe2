using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AUD_Gruppe2
{
    public  class WordWritter
    {
        Document doc;
        DocumentBuilder builder;
        Font myFont;

        public WordWriter(Document doc)
        {
            this.doc = doc;
            builder = new DocumentBuilder(doc);
            myFont = builder.Font;
            doc.Watermark.Remove();
        }
        public WordWriter()
        {
            doc = new Document();
            builder = new DocumentBuilder(doc);
            myFont = builder.Font;
            doc.Watermark.Remove();
        }

        public void changeFont(Font newFont)
        {
            //implement font change
        }

        public void AddText(string text)
        {
            builder.Writeln(text);
        }

        public void AddPicture(string path)
        {
            builder.InsertImage(path);
        }

        public void SaveToFile(string name)
        {
            doc.Save($"{name}.docx");
        }
    }
}
