using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace AUD_Gruppe2
{
    public  class WordWritter
    {
        Application wordApp = new Application();
        Document wordDoc;
        Range docRange;

        public WordWritter(Document doc)
        {
            this.wordDoc = doc;

        }
        public WordWritter()
        {
            this.wordDoc = wordApp.Documents.Add();
            docRange = wordDoc.Range();
        }

        public void AddText(string text)
        {

        }

        public void AddPicture(string path)
        {
            string imgPath = $"{Environment.CurrentDirectory}\\{path}";

            InlineShape autoScaledInlineShape = docRange.InlineShapes.AddPicture(imgPath);
            float scaledWidth = autoScaledInlineShape.Width;
            float scaledHeight = autoScaledInlineShape.Height;
            autoScaledInlineShape.Delete();

            // Create a new Shape and fill it with the picture
            Shape newShape = wordDoc.Shapes.AddShape(1, 0, 0, scaledWidth, scaledHeight);
            newShape.Fill.UserPicture(imgPath);
        }

        public void AddPictures(List<string> path)
        {

            int _counter = 0;
            float lastHeight = 0;
            int _spacer = 5;
            float _unitsOnPage = 0;


            int index = 1;
            foreach (string pathItem in path)
            {
                /*
                string imgPath1 = $"{Environment.CurrentDirectory}\\{pathItem}";

                InlineShape autoScaledInlineShape1 = docRange.InlineShapes.AddPicture(imgPath1);
                float scaledWidth1 = autoScaledInlineShape1.Width;
                float scaledHeight1 = autoScaledInlineShape1.Height;
                autoScaledInlineShape1.Delete();

                if ((scaledWidth1 + _spacer) + _unitsOnPage > 1000) {
                    docRange.InsertBreak(Word.WdBreakType.wdPageBreak);
                    _unitsOnPage = 0;
                    lastHeight = 0;
                    _counter++;   
                }

                // Create a new Shape and fill it with the picture
                Shape newShape1 = docRange.InlineShapes.AddShape(1, _counter, lastHeight, scaledWidth1, scaledWidth1);
                
                newShape1.Fill.UserPicture(imgPath1);
                lastHeight += scaledHeight1 + _spacer;
                _unitsOnPage += lastHeight;
                Console.WriteLine("Height of Picture : " + _counter + " : " + lastHeight);
                */
                //string imgPath1 = $"{Environment.CurrentDirectory}\\{pathItem}";
                var paragraph = wordDoc.Paragraphs.Add();
                paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphDistribute;
                paragraph.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                //paragraph.Range.InlineShapes.AddPicture(imgPath1);
                paragraph.Range.InlineShapes.AddPicture(pathItem);
                var paragraph2 = wordDoc.Paragraphs.Add();
                paragraph2.Range.Text = "Bild: "+index+"\r\r\r\r";
                Console.WriteLine("Pageindex of Picture : " + wordDoc.Paragraphs.Count + " : " + _counter);
                index++;
            }
        }

        public void SaveToFile(string name)
        {
            wordDoc.SaveAs2(name + ".docx");
            wordApp.Quit();
        }

    }
}
