namespace PDFAppend
{
    public struct PDFText
    {
        public string Text { get; set; }
        public float X { get; set; }
        public float Y { get; set; }
        public float BoxWidth { get; set; }
        public float BoxHeight { get; set; }
        public float FontSize { get; set; }
        public string Font { get; set; }
        public int Alignment { get; set; }
        public int Cyan { get; set; }
        public int Magenta { get; set; }
        public int Yellow { get; set; }
        public int Black { get; set; }
        public float Rotation { get; set; }
        public float Scale { get; set; }
        public string Orientation { get; set; }
        public float Leading { get; set; }

        public PDFText(string text, float x, float y, float boxWidth, float boxHeight, float fontSize, string font,
                       int alignment = 0, int cyan = 0, int magenta = 0, int yellow = 0, int black = 255,
                       float rotation = 0, float scale = 100f, float leading = 0, string orientation = PDFAppend.HORIZONTAL)
        {
            Text = text;
            X = x;
            Y = y;
            BoxWidth = boxWidth;
            BoxHeight = boxHeight;
            FontSize = fontSize;
            Font = font;
            Alignment = alignment;
            Cyan = cyan;
            Magenta = magenta;
            Yellow = yellow;
            Black = black;
            Rotation = rotation;
            Scale = scale;
            Leading = leading;
            Orientation = orientation;
        }
    }
}
