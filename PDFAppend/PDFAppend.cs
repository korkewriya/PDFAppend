using System;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using CSharp.Japanese.Kanaxs;

namespace PDFAppend
{
    public class PDFAppend
    {
        private string tmppath;

        public static PdfReader pdfReader { get; set; } 
        public static Document document { get; set; } 
        public static FileStream fs { get; set; } 
        public static PdfWriter writer { get; set; }

        public const int DEFAULT = 0;
        public const int TOP = 1;
        public const int RIGHT = 2;
        public const int LEFT = 3;
        public const int BOTTOM = 4;
        public const int CENTER = 5;

        public const string HORIZONTAL = "HORIZONTAL";
        public const string VERTICAL = "VERTICAL";

        // コンストラクタ::必要なファイルを開く
        public PDFAppend(string srcPDF)
        {
            tmppath = Path.GetTempFileName();

            pdfReader = new PdfReader(srcPDF);
            var size = pdfReader.GetPageSize(1);
            document = new Document(size);
            fs = new FileStream(this.tmppath, FileMode.Create, FileAccess.Write);
            writer = PdfWriter.GetInstance(document, fs);
            document.Open();
        }

        // インスタンスを閉じる
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
            File.Move(tmppath, dstpath);
        }

        // 一時ファイルのパスを返す
        public string GetTempPathName()
        {
            return tmppath;
        }

        // テンプレートをコピーする
        public PdfContentByte CopyTemplate()
        {
            var pdfContentByte = writer.DirectContent;
            var page = writer.GetImportedPage(pdfReader, 1);
            pdfContentByte.AddTemplate(page, 0, 0);
            return pdfContentByte;
        }

        // PDF画像を単純に表示する
        public PdfContentByte CopyPDFImg(PdfContentByte pdfContentByte, string PDFImg, float x, float y)
        {
            return PDFImage.CopyPDFImg(pdfContentByte, PDFImg, x, y);
        }

        // PDF画像をフレームいっぱいに収まるように表示する
        public PdfContentByte CopyPDFImgScaled(PdfContentByte pdfContentByte, string PDFImg, float x, float y, float boxWidth, float boxHeight)
        {
            return PDFImage.CopyPDFImgScaled(pdfContentByte, PDFImg, x, y, boxWidth, boxHeight);
        }

        // 画像をフレームいっぱいに表示する
        public PdfContentByte AddClippedImage(PdfContentByte pdfContentByte, string imagePath, float x, float y, float frmW, float frmH, int align = DEFAULT)
        {
            return PDFImage.AddClippedImage(pdfContentByte, imagePath, x, y, frmW, frmH, align);
        }

        // 矩形を描く
        public void MakeRectangle(PdfContentByte pdfContentByte, float x, float y, float boxWidth, float boxHeight,
                                  int cyan = 0, int magenta = 0, int yellow = 0, int black = 255, float opacity = 1.0f)
        {
            PDFImage.MakeRectangle(pdfContentByte, x, y, boxWidth, boxHeight, cyan, magenta, yellow, black, opacity);
        }

        // 1行テキストの追加
        public float Append(PdfContentByte pdfContentByte, string text, float x, float y, float boxWidth, float boxHeight,
                            float fontSize, string font, int alignment = 0, int cyan = 0, int magenta = 0, int yellow = 0, int black = 255,
                            float rotation = 0, float scale = 100f, string orientation = HORIZONTAL)
        {
            if (orientation == HORIZONTAL)
            {
                float ScaledWidth = HorizontalText.SetFontH(pdfContentByte, font, fontSize, text, boxWidth, cyan, magenta, yellow, black, scale);
                HorizontalText.ShowTextH(pdfContentByte, x, y, text, alignment, rotation);
                return ScaledWidth + x;
            }
            else
            {
                VerticalText.SetFontV(pdfContentByte, font, fontSize, text, cyan, magenta, yellow, black);
                pdfContentByte.SetHorizontalScaling(100f);
                float alignedY = VerticalText.ShowTextV(pdfContentByte, Kana.ToZenkaku(text), x, y, fontSize, boxHeight, alignment);
                return alignedY;
            }
        }

        /*
        public float Append(PdfContentByte pdfContentByte, PDFText pdfText)
        {
            return Append(pdfContentByte, pdfText.Text, pdfText.X, pdfText.Y, pdfText.BoxWidth, pdfText.BoxHeight,
                          pdfText.FontSize, pdfText.Font, pdfText.Alignment, pdfText.Cyan, pdfText.Magenta, pdfText.Yellow,
                          pdfText.Black, pdfText.Rotation, pdfText.Scale, pdfText.Orientation);
        }
        */

        // 複数行テキストの追加
        /*
        public void AppendMultiLine(PdfContentByte pdfContentByte, PDFText pdfText, int lines = 1, float lastPadding = 0)
        {
            AppendMultiLine(pdfContentByte, pdfText.Font, pdfText.FontSize, pdfText.Text, pdfText.X, pdfText.Y, pdfText.BoxWidth, pdfText.BoxHeight,
                            pdfText.Scale, pdfText.Cyan, pdfText.Magenta, pdfText.Yellow, pdfText.Black, pdfText.Leading, lines, lastPadding);
        }
        */

        public void AppendMultiLine(PdfContentByte pdfContentByte, string font, float fontSize, string text, float x, float y,
                                    float boxWidth, float boxHeight, float scale = 100f, int cyan = 0, int magenta = 0, int yellow = 0, int black = 255,
                                    float leading = 0, int lines = 1, float lastPadding = 0)
        {
            HorizontalText.AppendMultiLine(pdfContentByte, font, fontSize, text, x, y, boxWidth, boxHeight, scale, cyan, magenta, yellow, black,
                                           leading, lines, lastPadding);
        }

        // 合成フォント風テキストの追加
        /*
        public void AppendCompositeFontText(PdfContentByte pdfContentByte, PDFText pdfText, string compositeText, float fontSize2, string font2)
        {
            AppendCompositeFontText(pdfContentByte, pdfText.Text, compositeText, pdfText.X, pdfText.Y, pdfText.BoxWidth, pdfText.BoxHeight,
                                    pdfText.FontSize, pdfText.Font, fontSize2, font2, pdfText.Alignment, pdfText.Cyan, pdfText.Magenta, pdfText.Yellow, pdfText.Black,
                                    pdfText.Rotation, pdfText.Scale);
        }
        */

        public void AppendCompositeFontText(PdfContentByte pdfContentByte, string text, string compositeText, float x, float y,
                                            float boxWidth, float boxHeight, float fontSize, string font, float fontSize2, string font2,
                                            int alignment = 0, int cyan = 0, int magenta = 0, int yellow = 0, int black = 255, float rotation = 0, float scale = 100f)
        {
            HorizontalText.AppendCompositeFontText(pdfContentByte, text, compositeText, x, y, boxWidth, boxHeight, fontSize, font, fontSize2, font2,
                                                   alignment, cyan, magenta, yellow, black, rotation, scale);
        }

        // 複数行テキストの追加（改行で文字列分割）
        public void AppendMultiLineComment(PdfContentByte pdfContentByte, string font, float fontSize, string text, float x, float y, float boxWidth, float boxHeight,
                                           int lines, float leading, string orientation = HORIZONTAL, float scale = 100f,
                                           int cyan = 0, int magenta = 0, int yellow = 0, int black = 255, int alignment = Element.ALIGN_LEFT)
        {
            var bf = BaseFont.CreateFont(font, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            pdfContentByte.SetFontAndSize(bf, fontSize);
            pdfContentByte.SetHorizontalScaling(scale);

            pdfContentByte.SetCMYKColorFill(cyan, magenta, yellow, black);

            char[] separator = new char[] { '\r', '\n' };
            string[] comments = text.Split(separator, StringSplitOptions.RemoveEmptyEntries);

            // 横書きのパターン
            if(orientation == HORIZONTAL) {
                HorizontalText.SetMinScale(pdfContentByte, bf, comments, fontSize, boxWidth, scale);
                // y = y + boxHeight - leading;

                for (int i = 0; comments.Length > i; i++)
                {
                    if (i >= lines)
                        break;

                    string comment = comments[i];
                    HorizontalText.ShowTextH(pdfContentByte, x, y, comment, alignment, 0);
                    y -= leading;
                }
            }
            // 縦書きのパターン
            else
            {
                float vScale = VerticalText.GetMinVerticalScaling(comments, fontSize, boxHeight);
                x = VerticalText.GetAlignedXV(x + boxWidth, fontSize) - leading;
                y -= fontSize * vScale;

                pdfContentByte.BeginText();
                for (int i = 0; comments.Length > i; i++)
                {
                    if (i >= lines)
                        break;

                    string comment = comments[i];
                    VerticalText.MakeVerticalLine(pdfContentByte, Kana.ToZenkaku(comment), x, y, fontSize, alignment, vScale, tatechuyoko: true, isShugo: true);
                    x -= leading;
                }
                pdfContentByte.EndText();
            }
        }

        // 1行テキストの変倍文字幅を返す
        public float GetScaledWidth(string text, float boxWidth, float boxHeight, float fontSize, string font, float scale = 100f)
        {
            return HorizontalText.GetScaledWidth(text, boxWidth, boxHeight, fontSize, font, scale);
        }
    }
}
