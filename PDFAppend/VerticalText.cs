using System;
using iTextSharp.text;
using iTextSharp.text.pdf;
using CSharp.Japanese.Kanaxs;

namespace PDFAppend
{
    static class VerticalText
    {
        //　フォント・スケールを設定する
        public static void SetFontV(PdfContentByte pdfContentByte, string fontname, float fontsize, string text, int c, int m, int y, int k)
        {
            var bf = BaseFont.CreateFont(fontname, BaseFont.IDENTITY_V, BaseFont.EMBEDDED);
            pdfContentByte.SetFontAndSize(bf, fontsize);
            pdfContentByte.SetCMYKColorFill(c, m, y, k);
        }

        // PDFに文字情報を書き込む
        public static float ShowTextV(PdfContentByte pdfContentByte, string text, float x, float y, float fontSize, float boxHeight, int alignment, bool tatechuyoko = false)
        {
            pdfContentByte.BeginText();
            float vScale = GetVerticalScaling(text, fontSize, boxHeight);
            float VerticalPoint = GetVerticalPoint(text, fontSize);
            y = GetAlignedYV(alignment, y, boxHeight, VerticalPoint);
            x = GetAlignedXV(x, fontSize);

            y = MakeVerticalLine(pdfContentByte, text, x, y, fontSize, alignment, vScale);
            pdfContentByte.EndText();

            return y;
        }

        // 一行分の縦書きテキストを出力する
        public static float MakeVerticalLine(PdfContentByte pdfContentByte, string text, float x, float y, float fontSize, int alignment, float vScale,
                                      bool tatechuyoko = false, bool isShugo = false)
        {
            // index番目の文字を取得
            char GetChar(string str, int index)
            {
                try
                {
                    return str[index];
                }
                catch (IndexOutOfRangeException)
                {
                    return '_';
                }
            }

            string zenNum = "０１２３４５６７８９";
            string kakkoStart = "（「【『＜［〈《〔｛";
            string kakkoEnd = "）」】』＞］〉》〕｝";
            string rotate = "ー～－：；〜";
            string touten = "、。.,．，";
            bool asHankaku = false;

            for (var i = 0; text.Length > i; i++)
            {
                char c = text[i];
                char prev, next, nextnext;

                if ((kakkoStart.IndexOf(c) >= 0 || kakkoEnd.IndexOf(c) >= 0 || rotate.IndexOf(c) >= 0) && !isShugo)  // 括弧は90度回転させて移動（非集合チラシ）
                {
                    pdfContentByte.ShowTextAligned(alignment, c.ToString(), 0, -1 * vScale, 1, 0, x + (fontSize / 2), y - (fontSize * vScale / 2), 0);
                }
                else if ((kakkoStart.IndexOf(c) >= 0 || kakkoEnd.IndexOf(c) >= 0 || rotate.IndexOf(c) >= 0) && isShugo)  // 括弧は90度回転させて移動（集合チラシ）
                {
                    pdfContentByte.ShowTextAligned(alignment, c.ToString(), 0, -1 * vScale, 1, 0, x + (fontSize / 8), y + (fontSize * vScale * 7 / 8), 0);
                }
                else if (touten.IndexOf(c) >= 0)  // 読点
                {
                    pdfContentByte.ShowTextAligned(alignment, c.ToString(), 1, 0, 0, vScale, x + (fontSize / 2), y + (fontSize * vScale / 2), 0);
                }
                else if (zenNum.IndexOf(c) >= 0 && tatechuyoko)   // 数字（縦中横の設定）
                {
                    prev = GetChar(text, i - 1);
                    next = GetChar(text, i + 1);
                    nextnext = GetChar(text, i + 2);

                    if (zenNum.IndexOf(prev) < 0 && zenNum.IndexOf(next) >= 0 && zenNum.IndexOf(nextnext) < 0)
                    {
                        pdfContentByte.ShowTextAligned(alignment, Kana.ToHankaku(c.ToString()), 1, 0, 0, vScale, x - (fontSize / 6), y, 0);
                        pdfContentByte.ShowTextAligned(alignment, Kana.ToHankaku(next.ToString()), 1, 0, 0, vScale, x + (fontSize / 2), y, 0);
                        i++;
                    }
                    else
                    {
                        pdfContentByte.ShowTextAligned(alignment, c.ToString(), 1, 0, 0, vScale, x, y, 0);
                    }
                }
                else
                {
                    pdfContentByte.ShowTextAligned(alignment, c.ToString(), 1, 0, 0, vScale, x, y, 0);
                }

                // 半角文字かどうかを判定する
                next = GetChar(text, i + 1);
                if (kakkoEnd.IndexOf(c) >= 0 || kakkoStart.IndexOf(next) >= 0 || touten.IndexOf(c) >= 0)
                {
                    asHankaku = true;
                }

                if (asHankaku)
                {
                    y -= (fontSize * vScale) / 2.0f;
                }
                else
                {
                    y -= fontSize * vScale;
                }

                asHankaku = false;
            }

            return y;
        }

        // 指定幅に収まるよう長体をかけて調整する用の値を取得する
        public static float GetVerticalScaling(string text, float fontSize, float boxHeight)
        {
            float heightPoint = GetVerticalPoint(text, fontSize);
            if (boxHeight > heightPoint) return 1f;
            else return boxHeight / heightPoint;
        }

        // 縦書き文字の総ポイントを取得する
        public static float GetVerticalPoint(string text, float fontSize)
        {
            return text.Length * fontSize;
        }

        // alignを適用したY座標を取得する
        public static float GetAlignedYV(int alignment, float y, float boxHeight, float VerticalPoint)
        {
            if (VerticalPoint > boxHeight) VerticalPoint = boxHeight;

            if (alignment == Element.ALIGN_CENTER)
            {
                return y - (boxHeight - VerticalPoint) / 2.0f;
            }
            else if (alignment == Element.ALIGN_RIGHT)
            {
                return y - (boxHeight - VerticalPoint);
            }
            else
            {
                return y;
            }
        }

        // alignを適用したX座標を取得する
        public static float GetAlignedXV(float x, float fontSize)
        {
            return x; // + (fontSize / 2.0f);
        }

        // 複数テキストのうち、最小のスケール％を返す
        public static float GetMinVerticalScaling(string[] texts, float fontSize, float boxHeight)
        {
            float minScale = float.MaxValue;
            foreach (var text in texts)
            {
                float temp = GetVerticalScaling(text, fontSize, boxHeight);
                if (temp < minScale) minScale = temp;
            }
            return minScale;
        }
    }
}
