using System;
using System.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace PDFAppend
{
    static class HorizontalText
    {
        public static float GetScaledWidth(string text, float boxWidth, float boxHeight, float fontSize, string font, float scale = 100f)
        {
            var bf = BaseFont.CreateFont(font, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            float hScale = GetHorizontalScaling(bf, text, fontSize, boxWidth, scale);
            return GetWidthPointScaled(bf, text, fontSize, hScale);
        }

        // フォント・スケールを設定する
        public static float SetFontH(PdfContentByte pdfContentByte, string fontname, float fontsize, string text, float boxWidth, int c, int m, int y, int k, float scale)
        {
            var bf = BaseFont.CreateFont(fontname, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            pdfContentByte.SetFontAndSize(bf, fontsize);
            float hScale = GetHorizontalScaling(bf, text, fontsize, boxWidth, scale);
            pdfContentByte.SetHorizontalScaling(hScale);
            pdfContentByte.SetCMYKColorFill(c, m, y, k);
            return GetWidthPointScaled(bf, text, fontsize, hScale);
        }

        // PDFに文字情報を書き込む
        public static void ShowTextH(PdfContentByte pdfContentByte, float x, float y, string text, int alignment = Element.ALIGN_LEFT, float rotation = 0)
        {
            pdfContentByte.BeginText();
            pdfContentByte.ShowTextAligned(alignment, text, x, y, rotation);
            pdfContentByte.EndText();
        }

        // 指定幅に収まるよう長体をかけて調整する用の値を取得する
        public static float GetHorizontalScaling(BaseFont bf, string text, float fontsize, float boxWidth, float scale = 100f)
        {
            var widthPoint = bf.GetWidthPoint(text, fontsize);
            if (boxWidth > widthPoint) return scale;
            else return boxWidth / widthPoint * 100;
        }

        // 生成した文字の幅を取得する
        public static float GetWidthPointScaled(BaseFont bf, string text, float fontSize, float hScale)
        {
            var widthPoint = bf.GetWidthPoint(text, fontSize);
            return widthPoint * hScale / 100;
        }

        // 複数行テキストの追加
        public static void AppendMultiLine(PdfContentByte pdfContentByte, string font, float fontSize, string text, float x, float y, float boxWidth, float boxHeight,
                                           float scale, int cyan, int magenta, int yellow, int black, float leading, int lines, float lastPadding)
        {
            var bf = BaseFont.CreateFont(font, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            pdfContentByte.SetFontAndSize(bf, fontSize);
            pdfContentByte.SetCMYKColorFill(cyan, magenta, yellow, black);

            int maxLines = lines;

            // 仮の行数を取得
            lines = PredictMaxLine(bf, text, fontSize, boxWidth, lines);

            // 仮の長体%を取得（テキストを分割するのに使用）
            float lineHScale = GetHorizontalScaling(bf, text, fontSize, boxWidth - (lastPadding / lines), scale) * lines;
            if (lineHScale > 100f) lineHScale = 100f;

            var lineText = SplitTextToLine(text, bf, fontSize, lineHScale, lines, boxWidth);
            // lineText.Reverse();

            lineHScale = GetHorizontalScaling(bf, text, fontSize, boxWidth - (lastPadding / lineText.Count), scale) * lineText.Count;
            if (lineHScale > 100f)
            {
                lineHScale = 100f;
            }
            pdfContentByte.SetHorizontalScaling(lineHScale);

            // 最終行に空行を追加するかどうか決定する
            bool hasLastBlankLine = false;
            if(maxLines > lineText.Count)
            {
                string lastLine = lineText[lineText.Count - 1];
                float adjustedBoxWidth = boxWidth - lastPadding;
                float widthPoint = bf.GetWidthPoint(lastLine, fontSize) * lineHScale / 100f;
                float charSpace = 0.0f;
                float newWidthPoint = widthPoint;
                while (adjustedBoxWidth < newWidthPoint) // 枠幅 < 文字幅である限り
                {
                    charSpace -= 0.01f;
                    if (charSpace <= -0.5f)
                    {
                        lineText.Add("");
                        hasLastBlankLine = true;
                        break;
                    }
                }
            }

            lines = lineText.Count;

            int alignment;
            float startY = y + (leading * (lines - 1));
            for (var i = 0; lines > i; i++)
            {
                var line = lineText[i];

                if (i == lines - 1)
                {
                    alignment = Element.ALIGN_LEFT;
                    float adjustedBoxWidth = boxWidth - lastPadding;
                    float widthPoint = bf.GetWidthPoint(line, fontSize) * lineHScale / 100f;
                    float charSpace = 0.0f;
                    float newWidthPoint = widthPoint;
                    while (adjustedBoxWidth < newWidthPoint) // 枠幅 < 文字幅である限り
                    {
                        charSpace -= 0.01f;
                        newWidthPoint = widthPoint + (charSpace * lineHScale / 100f * (line.Length - 1));
                    }
                    pdfContentByte.SetCharacterSpacing(charSpace);  // 概要と＜仲介＞がぶつからないようにする値
                }
                else
                {
                    alignment = Element.ALIGN_LEFT;
                    var widthPoint = bf.GetWidthPoint(line, fontSize);
                    float charSpace = 0.0f;
                    if (hasLastBlankLine && i == lines - 2)  // 空行あり + 空行を除いた最終行の場合
                    {
                    }
                    else
                    {
                        charSpace = ((boxWidth - widthPoint * lineHScale / 100f) / (line.Length - 1)) * 100f / lineHScale;
                    }
                    pdfContentByte.SetCharacterSpacing(charSpace);
                }

                ShowTextH(pdfContentByte, x, startY, line, alignment, 0);
                pdfContentByte.SetCharacterSpacing(0.0f);   // 値を元に戻す
                startY -= leading;
            }
        }

        // テキストの最高行数を予測する
        private static int PredictMaxLine(BaseFont bf, string text, float fontSize, float boxWidth, int maxLine)
        {
            var widthPoint = bf.GetWidthPoint(text, fontSize);
            int predicted = (int)Math.Ceiling(widthPoint / boxWidth);
            if (predicted < maxLine) return predicted;
            return maxLine;
        }

        // テキストを複数行テキストに変換する
        private static List<string> SplitTextToLine(string text, BaseFont bf, float fontSize, float lineHScale, int lines, float boxWidth)
        {
            List<string> lineText = new List<string>();
            System.Text.StringBuilder[] tempLineText = new System.Text.StringBuilder[lines + 1];

            char[] kinsokuHead = new char[] { '、', '。', '）', ')', '.', ',' };
            char[] kinsokuLast = new char[] { '「', '（', '(', '■', '□' };

            for (var i = 0; lines >= i; i++)
            {
                System.Text.StringBuilder sb = new System.Text.StringBuilder("");
                tempLineText[i] = sb;
            }

            for (int chidx = 0, index = 0; text.Length > chidx; chidx++)
            {
                char c = text[chidx];
                char next;
                try
                {
                    next = text[chidx + 1];
                }
                catch (IndexOutOfRangeException)
                {
                    next = '\0';
                }

                tempLineText[index].Append(c);
                var widthPoint = bf.GetWidthPoint(tempLineText[index].ToString(), fontSize) * lineHScale / 100f;

                if (Array.IndexOf(kinsokuHead, next) != -1) continue;
                if (Array.IndexOf(kinsokuLast, c) != -1) continue;

                if ((c == '\r' || c == '\n') && index < lines - 1)
                {
                    index++;
                    if (c == '\r' && next == '\n') chidx++;
                }
                else if (widthPoint > boxWidth && index < lines - 1) index++;
            }

            foreach (var line in tempLineText)
            {
                if (line.ToString() == "") break;
                lineText.Add(line.ToString());
            }
            return lineText;
        }

        // 複数テキストのうち、最小のスケール％を返す
        public static void SetMinScale(PdfContentByte pdfContentByte, BaseFont bf, string[] texts, float fontSize, float boxWidth, float scale)
        {
            float minScale = float.MaxValue;
            foreach (var text in texts)
            {
                var widthPoint = bf.GetWidthPoint(text, fontSize);
                var temp = GetHorizontalScaling(bf, text, fontSize, boxWidth, scale);
                if (temp < minScale) minScale = temp;
            }
            pdfContentByte.SetHorizontalScaling(minScale);
        }

        // 合成フォント風テキストの追加
        public static void AppendCompositeFontText(PdfContentByte pdfContentByte, string text, string compositeStr,
                                                   float x, float y, float boxWidth, float boxHeight, float fontSize, string font, float fontSize2, string font2,
                                                   int alignment, int cyan, int magenta, int yellow, int black, float rotation, float scale)
        {
            string[] StringToArray(string str)
            {
                Regex regex = new Regex("(" + compositeStr + ")");  // 分割したい文字列を|で繋いだ形式であること
                var substrings = regex.Split(str).Where(s => s != String.Empty);
                return substrings.ToArray();
            }

            var strCheck = compositeStr.Split('|');
            var strArray = StringToArray(text);
            float hScale = 100;
            float allWidth = 0;

            // 複数フォントを適用した文字列幅を取得する
            for (var i = 0; strArray.Length > i; i++)
            {
                string s = strArray[i];

                // compositeStrに該当する文字の場合
                if (strCheck.Contains(s))
                {
                    allWidth += GetScaledWidth(s, boxWidth, boxHeight, fontSize, font, scale);
                }
                else
                {
                    allWidth += GetScaledWidth(s, boxWidth, boxHeight, fontSize2, font2, scale);
                }
            }

            hScale = (boxWidth / allWidth) * 100 > 100 ? 100 : (boxWidth / allWidth) * 100;

            for (var i = 0; strArray.Length > i; i++)
            {
                string s = strArray[i];

                // compositeStrに該当する文字の場合
                if (strCheck.Contains(s))
                {
                    float ScaledWidth = SetFontH(pdfContentByte, font, fontSize, s, boxWidth, cyan, magenta, yellow, black, hScale);
                    ShowTextH(pdfContentByte, x, y, s, alignment, rotation);
                    x += ScaledWidth;
                }
                else
                {
                    float ScaledWidth = SetFontH(pdfContentByte, font2, fontSize2, s, boxWidth, cyan, magenta, yellow, black, hScale);
                    ShowTextH(pdfContentByte, x, y, s, alignment, rotation);
                    x += ScaledWidth;
                }
            }
        }
    }
}
