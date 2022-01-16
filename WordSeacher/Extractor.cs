
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace WordSeacher
{
    class Extractor
    {
        ComObject<Excel.Application> comExcel;
        ComObject<Excel.Workbooks> comBooks;

        ComObject<Word.Application> comWord;
        ComObject<Word.Documents> comDocuments;

        class ComObject<T> : IDisposable
        {
            public T obj { get; }
            public string id { get; set; }
            //コンストラクタ
            public ComObject(T obj, string id)
            {
                this.obj = obj;
                this.id = id;
            }
            protected virtual void Dispose(bool disposing)
            {
                if (disposing)
                {
                    Marshal.ReleaseComObject(obj);
                    //Debug.Print("** Released: " + id);
                }
            }
            ~ComObject()
            {
                Dispose(false);
            }
            public void Dispose()
            {
                Dispose(true);
            }
        }

        //Excel抽出器の作成
        public void createExcel()
        {
            comExcel = new ComObject<Excel.Application>(
                new Excel.Application()
                {
                    Visible = false,
                    DisplayAlerts = false,
                    EnableEvents = false,
                }, "Excel.Application");
            comBooks = new ComObject<Excel.Workbooks>(
                comExcel.obj.Workbooks, "Excel.Workbooks");
        }
        //Word抽出器の作成
        public void createWord()
        {
            comWord = new ComObject<Word.Application>(
                new Word.Application()
                {
                    Visible = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
                }, "Word.Application");
            
            comDocuments = new ComObject<Word.Documents>(
                comWord.obj.Documents, "Word.Documents");

        }
        //Word本文抽出
        public string textOfWord(string path)
        {
            using (var comDocument = new ComObject<Word.Document>(
                comDocuments.obj.Open(path,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false), "Word.Document" + " : " + path))
            {
                var text = "";
                var tempFileList = new List<string>();

                try
                {
                    //本文
                    var tempFile = Path.GetTempFileName();
                    comDocument.obj.SaveAs(tempFile, FileFormat: Word.WdSaveFormat.wdFormatText);
                    using (var stream = new FileStream(tempFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (var reader = new StreamReader(stream, Encoding.GetEncoding("shift_jis")))
                    {
                        tempFileList.Add(tempFile);
                        text += reader.ReadToEnd();
                    }
                    //図形の中のテキスト
                    foreach (Word.Shape shape in comDocument.obj.Shapes)
                    {
                        ExtractText(shape, ref text);
                    }
                    return text;
                }
                finally
                {
                    comDocument.obj.Close();
                    foreach (string tempFile in tempFileList)
                    {
                        File.Delete(tempFile);
                    }
                }
            }
        }
        //Word図形テキスト抽出
        private void ExtractText(Word.Shape shape, ref string text)
        {
            shape.Select();

            comWord.obj.Visible = false;

            try
            {
                if (shape.Type == MsoShapeType.msoPicture ||
                               shape.Type == MsoShapeType.msoLine ||
                               shape.Type == MsoShapeType.msoAutoShape ||
                               shape.Type == MsoShapeType.msoFreeform ||
                               shape.Type == MsoShapeType.msoChart ||
                               (int)shape.Type == 28) //msoGraphic
                {
                    //nop
                }
                else if (shape.Type == MsoShapeType.msoGroup)
                {
                    foreach (Word.Shape child in shape.GroupItems)
                    {
                        ExtractText(child, ref text);
                    }
                }
                else
                {
                    if (shape.TextFrame != null && shape.TextFrame.HasText != 0)
                    {
                        var textInShape = shape.TextFrame?.TextRange?.Text;
                        if (!String.IsNullOrEmpty(textInShape))
                        {
                            text += textInShape;
                        }
                    }
                }
            }catch (Exception e)
            {
                Debug.Print(shape.Type.ToString() + " : " + e.Message);
            }
        }

        //Excel本文抽出
        public List<(string, string)> textOfExcel(string path)
        {
            using (var bookObj = new ComObject<Excel.Workbook>(
                    comBooks.obj.Open(
                    path,
                    UpdateLinks: Excel.XlUpdateLinks.xlUpdateLinksNever,
                    ReadOnly: true,
                    IgnoreReadOnlyRecommended: true,
                    Editable: false), "Excel.Workbook" + " : " + path))
            {
                var recordList = new List<(string, string)>();
                var tempFileList = new List<string>();
                try
                {
                    for (int i = 1; i <= bookObj.obj.Worksheets.Count; i++)
                    {
                        var text = "";
                        using (var comSheet = new ComObject<Excel.Worksheet>(bookObj.obj.Worksheets[i], 
                            "Excel.Worksheet" + "  " + bookObj.obj.Worksheets[i].Name))
                        {
                            var record = (comSheet.obj.Name, "");
                            //シート本体のテキスト
                            var tempFile = Path.GetTempFileName();
                            comSheet.obj.SaveAs(tempFile, FileFormat: Excel.XlFileFormat.xlCSV);
                            using (var stream = new FileStream(tempFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                            using (var reader = new StreamReader(stream, Encoding.GetEncoding("shift_jis")))
                            {
                                tempFileList.Add(tempFile);
                                text += reader.ReadToEnd();
                            }
                            //図形の中のテキスト
                            foreach (Excel.Shape shape in comSheet.obj.Shapes)
                            {
                                ExtractText(shape, ref text);
                            }

                            record.Item2 = text;
                            recordList.Add(record);
                        }
                    }
                    return recordList;
                }
                finally
                {
                    comBooks.obj.Close();
                    foreach (string tempFile in tempFileList)
                    {
                        File.Delete(tempFile);
                    }
                }
            }
        }
        //Excel図形テキスト抽出
        private void ExtractText(Excel.Shape shape, ref string text)
        {
            try
            {
                if (shape.Type == MsoShapeType.msoPicture ||
                    shape.Type == MsoShapeType.msoLine ||
                    shape.Type == MsoShapeType.msoAutoShape ||
                    shape.Type == MsoShapeType.msoFreeform ||
                    shape.Type == MsoShapeType.msoChart ||
                    (int)shape.Type == 28) //msoGraphic
                {
                    //nop
                }
                else if (shape.Type == MsoShapeType.msoGroup)
                {
                    foreach (Excel.Shape child in shape.GroupItems)
                    {
                        ExtractText(child, ref text);
                    }
                }
                else
                {
                    var textInShape = shape.TextFrame?.Characters()?.Text;
                    if (!String.IsNullOrEmpty(textInShape))
                    {
                        text += textInShape;
                    }
                }
            }
            catch(Exception e)
            {
                Debug.Print(shape.Type.ToString() + " : " + e.Message);
            }
        }

        //後処理
        public void CleanUp()
        {
            if (comExcel != null)
            {
                comExcel.obj.Quit();
                comBooks.Dispose();
                comExcel.Dispose();
            }
            if (comWord != null)
            {
                comWord.obj.Quit();
                comDocuments.Dispose();
                comWord.Dispose();
            }
        }
    }
}
