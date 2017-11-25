using System;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace PDFAppender
{
    public class PDFAppend
    {
        private string srcpath;
        private string tmppath;

        private PdfReader pdfReader;
        private Document document;
        private FileStream fs;
        private PdfWriter writer;

        // コンストラクタ::必要なファイルを開く
        public PDFAppend(string srcPDF)
        {
            this.tmppath = Path.GetTempFileName();

            pdfReader = new PdfReader(srcPDF);
            var size = pdfReader.GetPageSize(1);
            document = new Document(size);
            fs = new FileStream(this.tmppath, FileMode.Create, FileAccess.Write);
            writer = PdfWriter.GetInstance(document, fs);
            document.Open();
        }

        // テンプレートをコピーする
        public PdfContentByte CopyTemplate()
        {
            var pdfContentByte = writer.DirectContent;
            var page = writer.GetImportedPage(pdfReader, 1);
            pdfContentByte.AddTemplate(page, 0, 0);
            return pdfContentByte;
        }

        // PDF形式の画像をコピーする
        public PdfContentByte CopyPDFImg(PdfContentByte pdfContentByte, string PDFImg, float x, float y)
        {
            PdfReader img = new PdfReader(PDFImg);
            var page = writer.GetImportedPage(img, 1);
            pdfContentByte.AddTemplate(page, x, y);
            return pdfContentByte;
        }

        // 指定座標・フォントでテキストを追記する
        public void Append(ref PdfContentByte pdfContentByte, string str,
                          float x, float y, float boxWidth, float fontSize, string FONT, int alignment = 0,
                          int cyan = 0, int magenta = 0, int yellow = 0, int black = 100)
        {
            SetFont(ref pdfContentByte, FONT, fontSize, str, boxWidth, cyan, magenta, yellow, black);
            ShowTextAligned(pdfContentByte, x, y + 1, str, alignment);
        }

        // クラスを閉じる
        public void Close()
        {
            document.Close();
            fs.Close();
            writer.Close();
            pdfReader.Close();
        }

        // 一時ファイルに出力しているデータを移動させる
        public void Save(string saveDir, string dstPDF)
        {
            string dstpath = Path.Combine(saveDir, dstPDF);
            File.Move(this.tmppath, dstpath);
        }

        public string GetTempPathName()
        {
            return this.tmppath;
        }

        // PDFに文字情報を書き込む
        public void ShowTextAligned(PdfContentByte pdfContentByte, float x, float y, string text, int alignment = Element.ALIGN_LEFT, float rotation = 0)
        {
            pdfContentByte.BeginText();
            pdfContentByte.ShowTextAligned(alignment, text, x, y, rotation);
            pdfContentByte.EndText();
        }

        // フォント・スケールを設定する
        public void SetFont(ref PdfContentByte pdfContentByte, string fontname, float fontsize, string text, float boxWidth, int c, int m, int y, int k)
        {
            var bf = BaseFont.CreateFont(fontname, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            pdfContentByte.SetFontAndSize(bf, fontsize);
            float hScale = this.GetHorizontalScaling(bf, text, fontsize, boxWidth);
            pdfContentByte.SetHorizontalScaling(hScale);
            pdfContentByte.SetCMYKColorFill(c, m, y, k);
        }

        // 指定幅に収まるよう長体をかけて調整する用の値を取得する
        public float GetHorizontalScaling(BaseFont bf, string text, float fontsize, float boxWidth)
        {
            var widthPoint = bf.GetWidthPoint(text, fontsize);
            if (boxWidth > widthPoint) return 100f;
            else return boxWidth / widthPoint * 100;
        }
    }
}
