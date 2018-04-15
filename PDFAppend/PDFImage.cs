using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace PDFAppend
{
    static class PDFImage
    {
        // PDF形式の画像をコピーする
        public static PdfContentByte CopyPDFImg(PdfContentByte pdfContentByte, string PDFImg, float x, float y)
        {
            PdfReader img = new PdfReader(PDFImg);
            var page = PDFAppend.writer.GetImportedPage(img, 1);
            pdfContentByte.AddTemplate(page, x, y);
            return pdfContentByte;
        }

        // 画像PDFを等倍で枠いっぱいに表示させる
        public static PdfContentByte CopyPDFImgScaled(PdfContentByte pdfContentByte, string PDFImg, float x, float y, float boxWidth, float boxHeight)
        {
            PdfReader img = new PdfReader(PDFImg);
            var page = PDFAppend.writer.GetImportedPage(img, 1);

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
        public static PdfContentByte AddClippedImage(PdfContentByte pdfContentByte, string imagePath, float x, float y, float frmW, float frmH, int align)
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
                if ((align != PDFAppend.CENTER && (frmW / frmH > imgW / imgH)) || (align == PDFAppend.CENTER && (frmW / frmH < imgW / imgH)))
                {
                    float scaledWidth = frmW;
                    scale = frmW / imgW;
                    float scaledHeight = imgH * scale;

                    if (align == PDFAppend.TOP) y = y - (scaledHeight - frmH);
                    else if (align == PDFAppend.BOTTOM) DoNothing();
                    else y = y - ((scaledHeight - frmH) / 2);
                }
                else  // 画像横長の場合
                {
                    float scaledHeight = frmH;
                    scale = frmH / imgH;
                    float scaledWidth = imgW * scale;

                    if (align == PDFAppend.LEFT) DoNothing();
                    else if (align == PDFAppend.RIGHT) x = x - (scaledWidth - frmW);
                    else x = x - ((scaledWidth - frmW) / 2);
                }
                image.ScalePercent(scale * 100);
                image.SetAbsolutePosition(x, y);
                pdfContentByte.AddImage(image);
                pdfContentByte.RestoreState();
            }
            return pdfContentByte;
        }

        // 矩形を作成する
        public static void MakeRectangle(PdfContentByte pdfContentByte, float x, float y, float boxWidth, float boxHeight,
                                         int cyan = 0, int magenta = 0, int yellow = 0, int black = 255, float opacity = 1.0f)
        {
            pdfContentByte.Rectangle(x, y, boxWidth, boxHeight);
            pdfContentByte.SetCMYKColorFill(cyan, magenta, yellow, black);
            if (opacity != 1.0f)
            {
                SetOpacity(pdfContentByte, opacity);

            }
            pdfContentByte.Fill();
            // 透明度情報を戻す
            if (opacity != 1.0f)
            {
                SetOpacity(pdfContentByte, 1.0f);
            }
        }

        // オブジェクト透明度の指定
        public static void SetOpacity(PdfContentByte pdfContentByte, float opacity)
        {
            PdfGState graphicsState = new PdfGState();
            graphicsState.FillOpacity = opacity;
            pdfContentByte.SetGState(graphicsState);
        }
    }
}
