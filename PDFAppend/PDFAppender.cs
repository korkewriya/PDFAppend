using System;
using System.IO;
using System.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;
using CSharp.Japanese.Kanaxs;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace PDFAppender
{
    public class PDFAppend
    {
        private string tmppath;

        private PdfReader pdfReader;
        private Document document;
        private FileStream fs;
        private PdfWriter writer;

        private const int DEFAULT = 0;
        private const int TOP = 1;
        private const int RIGHT = 2;
        private const int LEFT = 3;
        private const int BOTTOM = 4;
        private const int CENTER = 5;

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

        // 画像PDFを等倍で枠いっぱいに表示させる
        public PdfContentByte CopyPDFImgScaled(PdfContentByte pdfContentByte, string PDFImg, float x, float y, float boxWidth, float boxHeight)
        {
            PdfReader img = new PdfReader(PDFImg);
            var page = writer.GetImportedPage(img, 1);

            float height = img.GetPageSize(1).Height;
            float width = img.GetPageSize(1).Width;

            if (boxWidth / boxHeight > width / height)   // 画像を高さいっぱいに配置
            {
                float scaledHeight = boxHeight;
                float scaledFactor = boxHeight / height;
                float scaledWidth = width * scaledFactor;
                float offset = (boxWidth - scaledWidth) / 2;
                pdfContentByte.AddTemplate(page, scaledFactor, 0, 0, scaledFactor, x + offset, y);
            }
            else    // 画像を幅いっぱいに配置
            {
                float scaledWidth = boxWidth;
                float scaledFactor = boxWidth / width;
                float scaledHeight = height * scaledFactor;
                float offset = (boxHeight - scaledHeight) / 2;
                pdfContentByte.AddTemplate(page, scaledFactor, 0, 0, scaledFactor, x, y + offset);
            }

            return pdfContentByte;
        }

        // 画像をクリッピングして貼り付ける
        public PdfContentByte AddClippedImage(PdfContentByte pdfContentByte, string imagePath, float x, float y, float frmW, float frmH, int align = DEFAULT)
        {
            void DoNothing() { }    // 何もしない関数を定義

            using (Stream inputImageStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                Image image = Image.GetInstance(inputImageStream);
                float imgW = image.ScaledWidth;
                float imgH = image.ScaledHeight;
                float scale;

                pdfContentByte.SaveState();   // クリッピングパス削除のため使用
                pdfContentByte.Rectangle(x, y, frmW, frmH);
                pdfContentByte.Clip();
                pdfContentByte.NewPath();

                // 画像縦長の場合
                if ((align != CENTER && (frmW / frmH > imgW / imgH)) || (align == CENTER && (frmW / frmH < imgW / imgH)))
                {
                    float scaledWidth = frmW;
                    scale = frmW / imgW;
                    float scaledHeight = imgH * scale;

                    if (align == TOP) y = y - (scaledHeight - frmH);
                    else if (align == BOTTOM) DoNothing();
                    else y = y - ((scaledHeight - frmH) / 2);
                }
                else  // 画像横長の場合
                {
                    float scaledHeight = frmH;
                    scale = frmH / imgH;
                    float scaledWidth = imgW * scale;

                    if (align == LEFT) DoNothing();
                    else if(align == RIGHT) x = x - (scaledWidth - frmW);
                    else x = x - ((scaledWidth - frmW) / 2);
                }
                image.ScalePercent(scale * 100);
                image.SetAbsolutePosition(x, y);
                pdfContentByte.AddImage(image);
                pdfContentByte.RestoreState();
            }
            return pdfContentByte;
        }

        // 指定座標・フォントでテキストを追記する
        public float Append(ref PdfContentByte pdfContentByte, string str,
                            float x, float y, float boxWidth, float boxHeight, float fontSize, string FONT, int alignment = 0,
                            int cyan = 0, int magenta = 0, int yellow = 0, int black = 255, float rotation = 0, float scale = 100f,
                            string orientation = "HORIZONTAL")
        {
            if(orientation == "HORIZONTAL") { 
                float ScaledWidth = SetFontH(ref pdfContentByte, FONT, fontSize, str, boxWidth, cyan, magenta, yellow, black, scale);
                ShowTextH(pdfContentByte, x, y + 1, str, alignment, rotation);
                return ScaledWidth + x;
            }
            else
            {
                SetFontV(ref pdfContentByte, FONT, fontSize, str, cyan, magenta, yellow, black);
                pdfContentByte.SetHorizontalScaling(100f);
                float alignedY = ShowTextV(pdfContentByte, Kana.ToZenkaku(str), x, y, fontSize, boxHeight, alignment);
                return alignedY;
            }
        }

        public float GetScaledWidth(string str, float boxWidth, float boxHeight, float fontSize, string FONT, float scale = 100f)
        {
            var bf = BaseFont.CreateFont(FONT, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            float hScale = GetHorizontalScaling(bf, str, fontSize, boxWidth, scale);
            return GetWidthPointScaled(bf, str, fontSize, hScale);
        }

        // フォント・スケールを設定する（横）
        private float SetFontH(ref PdfContentByte pdfContentByte, string fontname, float fontsize, string text, float boxWidth, int c, int m, int y, int k, float scale)
        {
            var bf = BaseFont.CreateFont(fontname, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            pdfContentByte.SetFontAndSize(bf, fontsize);
            float hScale = GetHorizontalScaling(bf, text, fontsize, boxWidth, scale);
            pdfContentByte.SetHorizontalScaling(hScale);
            pdfContentByte.SetCMYKColorFill(c, m, y, k);
            return GetWidthPointScaled(bf, text, fontsize, hScale);
        }

        // PDFに文字情報を書き込む（横）
        private void ShowTextH(PdfContentByte pdfContentByte, float x, float y, string text, int alignment = Element.ALIGN_LEFT, float rotation = 0)
        {
            pdfContentByte.BeginText();
            pdfContentByte.ShowTextAligned(alignment, text, x, y, rotation);
            pdfContentByte.EndText();
        }

        // 指定幅に収まるよう長体をかけて調整する用の値を取得する（横）
        private float GetHorizontalScaling(BaseFont bf, string text, float fontsize, float boxWidth, float scale = 100f)
        {
            var widthPoint = bf.GetWidthPoint(text, fontsize);
            if (boxWidth > widthPoint) return scale;
            else return boxWidth / widthPoint * 100;
        }

        // 生成した文字の幅を取得する（横）
        private float GetWidthPointScaled(BaseFont bf, string text, float fontSize, float hScale)
        {
            var widthPoint = bf.GetWidthPoint(text, fontSize);
            return widthPoint * hScale / 100;
        }

        // PDFに文字情報を書き込む（縦）
        private float ShowTextV(PdfContentByte pdfContentByte, string text, float x, float y, float fontSize, float boxHeight, int alignment, bool tatechuyoko = false)
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
        private float MakeVerticalLine(PdfContentByte pdfContentByte, string text, float x, float y, float fontSize, int alignment, float vScale,
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
            string kakko = "（）「」【】『』＜＞［］〈〉《》〔〕｛｝ー～－";

            for (var i = 0; text.Length > i; i++)
            {
                char c = text[i];
                char prev, next, nextnext;

                if (kakko.IndexOf(c) >= 0 && !isShugo)  // 括弧は90度回転させて移動（非集合チラシ）
                {
                    pdfContentByte.ShowTextAligned(alignment, c.ToString(), 0, -1 * vScale, 1, 0, x + (fontSize / 2), y - (fontSize * vScale / 2), 0);
                }
                else if (kakko.IndexOf(c) >= 0 && isShugo)  // 括弧は90度回転させて移動（集合チラシ）
                {
                    pdfContentByte.ShowTextAligned(alignment, c.ToString(), 0, -1 * vScale, 1, 0, x + (fontSize / 8), y + (fontSize * vScale * 7 / 8), 0);
                }
                else if ("、。.,．，".IndexOf(c) >= 0)  // 読点
                {
                    pdfContentByte.ShowTextAligned(alignment, c.ToString(), 1, 0, 0, vScale, x + (fontSize / 2), y + (fontSize * vScale / 2), 0);
                }
                else if(zenNum.IndexOf(c) >= 0 && tatechuyoko)   // 数字（縦中横の設定）
                {
                    prev = GetChar(text, i - 1);
                    next = GetChar(text, i + 1);
                    nextnext = GetChar(text, i + 2);

                    if (zenNum.IndexOf(prev) < 0 && zenNum.IndexOf(next) >= 0 && zenNum.IndexOf(nextnext) < 0)
                    {
                        pdfContentByte.ShowTextAligned(alignment, Kana.ToHankaku(c.ToString()), 1, 0, 0, vScale, x - (fontSize / 8), y, 0);
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

                y -= fontSize * vScale;
            }

            return y;
        }

        // alignを適用したY座標を取得する
        private float GetAlignedYV(int alignment, float y, float boxHeight, float VerticalPoint)
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
        private float GetAlignedXV(float x, float fontSize)
        {
            return x; // + (fontSize / 2.0f);
        }

        //　フォント・スケールを設定する（縦）
        private void SetFontV(ref PdfContentByte pdfContentByte, string fontname, float fontsize, string text, int c, int m, int y, int k)
        {
            var bf = BaseFont.CreateFont(fontname, BaseFont.IDENTITY_V, BaseFont.EMBEDDED);
            pdfContentByte.SetFontAndSize(bf, fontsize);
            pdfContentByte.SetCMYKColorFill(c, m, y, k);
        }

        // 指定幅に収まるよう長体をかけて調整する用の値を取得する（縦）
        private float GetVerticalScaling(string text, float fontSize, float boxHeight)
        {
            float heightPoint = GetVerticalPoint(text, fontSize);
            if (boxHeight > heightPoint) return 1f;
            else return boxHeight / heightPoint;
        }

        // 縦書き文字の総ポイントを取得する
        private float GetVerticalPoint(string text, float fontSize)
        {
            return text.Length * fontSize;
        }

        // オブジェクト透明度の指定
        public void SetOpacity(PdfContentByte pdfContentByte, float opacity)
        {
            PdfGState graphicsState = new PdfGState();
            graphicsState.FillOpacity = opacity;
            pdfContentByte.SetGState(graphicsState);
        }

        // 複数行テキストの追加
        public void AppendMultiLine(PdfContentByte pdfContentByte, string fontname, float fontSize, string text, float x, float y, float boxWidth, float boxHeight,
                                    int lines, float leading, float lastPadding = 0, float scale = 100f, int cyan = 0, int magenta = 0, int yellow = 0, int keyplate = 255)
        {
            var bf = BaseFont.CreateFont(fontname, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            pdfContentByte.SetFontAndSize(bf, fontSize);
            pdfContentByte.SetCMYKColorFill(cyan, magenta, yellow, keyplate);

            // 仮の行数を取得
            lines = PredictMaxLine(bf, text, fontSize, boxWidth, lines);

            // 仮の長体%を取得（テキストを分割するのに使用）
            float lineHScale = GetHorizontalScaling(bf, text, fontSize, boxWidth - (lastPadding / lines), scale) * lines;
            if (lineHScale > 100f) lineHScale = 100f;

            var lineText = SplitTextToLine(text, bf, fontSize, lineHScale, lines, boxWidth);
            lineText.Reverse();

            lines = lineText.Count;
            lineHScale = GetHorizontalScaling(bf, text, fontSize, boxWidth - (lastPadding / lines), scale) * lines;
            if (lineHScale > 100f) lineHScale = 100f;
            pdfContentByte.SetHorizontalScaling(lineHScale);

            int alignment;
            for (var i = 0; lineText.Count > i; i++)
            {
                var line = lineText[i];

                if (i == 0)
                {
                    alignment = Element.ALIGN_LEFT;
                    float adjustedBoxWidth = boxWidth - lastPadding;
                    float widthPoint = bf.GetWidthPoint(line, fontSize) * lineHScale / 100f;
                    float charSpace = 0.0f;
                    while (adjustedBoxWidth < widthPoint) // 枠幅 < 文字幅である限り
                    {
                        charSpace -= 0.01f;
                        widthPoint += charSpace * lineHScale / 100f * (line.Length - 1);
                    }
                    pdfContentByte.SetCharacterSpacing(charSpace * 3);  // ??? 要検証
                }
                else
                {
                    alignment = Element.ALIGN_LEFT;
                    var widthPoint = bf.GetWidthPoint(line, fontSize);
                    float charSpace = ((boxWidth - widthPoint * lineHScale / 100f) / (line.Length - 1)) * 100f / lineHScale;

                    /*  デバッグ用：行幅を取得  */
                    /*
                    float total = 0;
                    for(var j = 0; line.Length > j; j++)
                    {
                        char c = line[j];
                        float charWidth = bf.GetWidthPoint(c, fontSize) * lineHScale / 100f;
                        total += charWidth;
                        if (j != line.Length - 1) total += charSpace;
                    }
                    */
                    pdfContentByte.SetCharacterSpacing(charSpace);
                }

                ShowTextH(pdfContentByte, x, y, line, alignment, 0);
                pdfContentByte.SetCharacterSpacing(0.0f);   // 値を元に戻す
                y += leading;
            }
        }

        // テキストを複数行テキストに変換する
        private List<string> SplitTextToLine(string text, BaseFont bf, float fontSize, float lineHScale, int lines, float boxWidth)
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

            for(int chidx = 0, index = 0; text.Length > chidx; chidx++)
            {
                char c = text[chidx];
                char next;
                try
                {
                    next = text[chidx + 1];
                }
                catch(IndexOutOfRangeException)
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

            foreach(var line in tempLineText)
            {
                if (line.ToString() == "") break;
                lineText.Add(line.ToString());
            }
            return lineText;
        }

        // 複数行テキストの追加（改行で文字列分割）
        public void AppendMultiLineComment(PdfContentByte pdfContentByte, string fontname, float fontSize, string text, float x, float y, float boxWidth, float boxHeight,
                                           int lines, float leading, string orientation = "HORIZONTAL", float scale = 100f,
                                           int cyan = 0, int magenta = 0, int yellow = 0, int keyplate = 255, int alignment = Element.ALIGN_LEFT)
        {
            var bf = BaseFont.CreateFont(fontname, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            pdfContentByte.SetFontAndSize(bf, fontSize);
            pdfContentByte.SetHorizontalScaling(scale);

            pdfContentByte.SetCMYKColorFill(cyan, magenta, yellow, keyplate);

            char[] separator = new char[] { '\r', '\n' };
            string[] comments = text.Split(separator, StringSplitOptions.RemoveEmptyEntries);

            // 横書きのパターン
            if(orientation == "HORIZONTAL") {
                SetMinScale(pdfContentByte, bf, comments, fontSize, boxWidth, scale);
                y = y + boxHeight - leading;

                for (int i = 0; comments.Length > i; i++)
                {
                    if (i >= lines) break;

                    string comment = comments[i];
                    ShowTextH(pdfContentByte, x, y, comment, alignment, 0);
                    y -= leading;
                }
            }
            // 縦書きのパターン
            else
            {
                float vScale = GetMinVerticalScaling(comments, fontSize, boxHeight);
                x = GetAlignedXV(x + boxWidth, fontSize) - leading;
                y -= fontSize * vScale;

                pdfContentByte.BeginText();
                for (int i = 0; comments.Length > i; i++)
                {
                    if (i >= lines) break;

                    string comment = comments[i];
                    MakeVerticalLine(pdfContentByte, Kana.ToZenkaku(comment), x, y, fontSize, alignment, vScale, tatechuyoko: true, isShugo: true);
                    x -= leading;
                }
                pdfContentByte.EndText();
            }
        }

        // 複数テキストのうち、最小のスケール％を返す（横）
        private void SetMinScale(PdfContentByte pdfContentByte, BaseFont bf, string[] texts, float fontSize, float boxWidth, float scale)
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

        // 複数テキストのうち、最小のスケール％を返す（縦）
        private float GetMinVerticalScaling(string[] texts, float fontSize, float boxHeight)
        {
            float minScale = float.MaxValue;
            foreach (var text in texts)
            {
                float temp = GetVerticalScaling(text, fontSize, boxHeight);
                if (temp < minScale) minScale = temp;
            }
            return minScale;
        }

        // テキストの最高行数を予測する
        private int PredictMaxLine(BaseFont bf, string text, float fontSize, float boxWidth, int maxLine)
        {
            var widthPoint = bf.GetWidthPoint(text, fontSize);
            int predicted = (int)Math.Ceiling(widthPoint / boxWidth);
            if (predicted < maxLine) return predicted;
            return maxLine;
        }

        // 矩形を作成する
        public void MakeRectangle(PdfContentByte pdfContentByte, float x, float y, float boxWidth, float boxHeight,
                                  int cyan = 0, int magenta = 0, int yellow = 0, int keyplate = 255, float opacity = 1.0f)
        {
            pdfContentByte.Rectangle(x, y, boxWidth, boxHeight);
            pdfContentByte.SetCMYKColorFill(cyan, magenta, yellow, keyplate);
            if (opacity != 1.0f)
                SetOpacity(pdfContentByte, opacity); 
            pdfContentByte.Fill();
            if (opacity != 1.0f)    // 透明度情報を戻す
                SetOpacity(pdfContentByte, 1.0f);
        }

        // 合成フォント風テキストの追加
        public void AppendCompositeFontText(ref PdfContentByte pdfContentByte, string str, string compositeStr,
                                            float x, float y, float boxWidth, float boxHeight, float fontSize, string FONT, float fontSize2, string FONT2,
                                            int alignment = 0, int cyan = 0, int magenta = 0, int yellow = 0, int black = 255, float rotation = 0,
                                            float scale = 100f, string orientation = "HORIZONTAL")
        {
            var strCheck = compositeStr.Split('|');
            var strArray = StringToArray(str, compositeStr);
            float hScale = 100;
            float allWidth = 0;

            // 複数フォントを適用した文字列幅を取得する
            for (var i = 0; strArray.Length > i; i++)
            {
                string s = strArray[i];

                // compositeStrに該当する文字の場合
                if (strCheck.Contains(s))
                {
                    allWidth += GetScaledWidth(s, boxWidth, boxHeight, fontSize, FONT, scale);
                }
                else
                {
                    allWidth += GetScaledWidth(s, boxWidth, boxHeight, fontSize2, FONT2, scale);
                }
            }

            hScale = (boxWidth / allWidth) * 100 > 100 ? 100 : (boxWidth / allWidth) * 100;

            for (var i = 0; strArray.Length > i; i++)
            {
                string s = strArray[i];

                // compositeStrに該当する文字の場合
                if (strCheck.Contains(s))
                {
                    float ScaledWidth = SetFontH(ref pdfContentByte, FONT, fontSize, s, boxWidth, cyan, magenta, yellow, black, hScale);
                    ShowTextH(pdfContentByte, x, y, s, alignment, rotation);
                    x += ScaledWidth;
                }
                else
                {
                    float ScaledWidth = SetFontH(ref pdfContentByte, FONT2, fontSize2, s, boxWidth, cyan, magenta, yellow, black, hScale);
                    ShowTextH(pdfContentByte, x, y, s, alignment, rotation);
                    x += ScaledWidth;
                }
            }
        }

        private string[] StringToArray(string str, string compositeStr)
        {
            Regex regex = new Regex("(" + compositeStr + ")");  // 分割したい文字列を|で繋いだ形式であること
            var substrings = regex.Split(str).Where(s => s != String.Empty);
            return substrings.ToArray();
        }
    }
}
