using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace LeeExcel
{
    public static class BrandIconHelper
    {
        private static Bitmap _cachedIcon32;

        public static Bitmap GetExcelMindIcon(int size = 32)
        {
            if (size == 32 && _cachedIcon32 != null)
            {
                return _cachedIcon32;
            }

            var bmp = new Bitmap(size, size, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                float scale = size / 32f;

                // 1. 经典 Excel 绿与科技感高饱和度渐变徽标 (Rounded Rect)
                using (var path = CreateRoundedRectanglePath(1 * scale, 1 * scale, (size - 2 * scale), (size - 2 * scale), 6 * scale))
                {
                    using (var brush = new LinearGradientBrush(
                        new PointF(0, 0),
                        new PointF(0, size),
                        ColorTranslator.FromHtml("#107C41"),
                        ColorTranslator.FromHtml("#094B26")))
                    {
                        g.FillPath(brush, path);
                    }

                    // 徽标微光边框
                    using (var pen = new Pen(ColorTranslator.FromHtml("#2EB86F"), 1.2f * scale))
                    {
                        g.DrawPath(pen, path);
                    }
                }

                // 2. 左下角精致 Excel 网格元素
                using (var gridPen = new Pen(Color.FromArgb(60, 255, 255, 255), 1.0f * scale))
                {
                    g.DrawLine(gridPen, 6 * scale, 21 * scale, 16 * scale, 21 * scale);
                    g.DrawLine(gridPen, 6 * scale, 25 * scale, 16 * scale, 25 * scale);
                    g.DrawLine(gridPen, 11 * scale, 18 * scale, 11 * scale, 27 * scale);
                }

                // 3. AI 神经网络连线 (Mind Synapses)
                float cx = 18.5f * scale;
                float cy = 14.0f * scale;

                using (var synPen = new Pen(ColorTranslator.FromHtml("#7EF6D4"), 1.5f * scale))
                {
                    synPen.StartCap = LineCap.Round;
                    synPen.EndCap = LineCap.Round;

                    g.DrawLine(synPen, cx, cy, 9 * scale, 9.5f * scale);   // 左上
                    g.DrawLine(synPen, cx, cy, 25 * scale, 9.5f * scale);  // 右上
                    g.DrawLine(synPen, cx, cy, 24 * scale, 21.5f * scale); // 右下
                    g.DrawLine(synPen, cx, cy, 12 * scale, 18.5f * scale); // 左下
                }

                // 4. 神经元发光节点 (Neural Nodes)
                using (var nodeBrush = new SolidBrush(ColorTranslator.FromHtml("#A8FFEB")))
                {
                    float nr = 2.0f * scale;
                    g.FillEllipse(nodeBrush, 9 * scale - nr, 9.5f * scale - nr, nr * 2, nr * 2);
                    g.FillEllipse(nodeBrush, 25 * scale - nr, 9.5f * scale - nr, nr * 2, nr * 2);
                    g.FillEllipse(nodeBrush, 24 * scale - nr, 21.5f * scale - nr, nr * 2, nr * 2);
                    g.FillEllipse(nodeBrush, 12 * scale - nr, 18.5f * scale - nr, nr * 2, nr * 2);
                }

                // 5. 核心 AI 智能星芒 (Sparkle Mind Core)
                using (var sparklePath = new GraphicsPath())
                {
                    float rOut = 6.2f * scale;
                    float rIn = 1.6f * scale;
                    PointF[] pts = new PointF[]
                    {
                        new PointF(cx, cy - rOut),
                        new PointF(cx + rIn, cy - rIn),
                        new PointF(cx + rOut, cy),
                        new PointF(cx + rIn, cy + rIn),
                        new PointF(cx, cy + rOut),
                        new PointF(cx - rIn, cy + rIn),
                        new PointF(cx - rOut, cy),
                        new PointF(cx - rIn, cy - rIn)
                    };
                    sparklePath.AddPolygon(pts);

                    using (var spBrush = new SolidBrush(Color.White))
                    {
                        g.FillPath(spBrush, sparklePath);
                    }
                }

                // 中心璀璨光核
                using (var coreBrush = new SolidBrush(ColorTranslator.FromHtml("#E8FFFF")))
                {
                    g.FillEllipse(coreBrush, cx - 1.0f * scale, cy - 1.0f * scale, 2.0f * scale, 2.0f * scale);
                }
            }

            if (size == 32)
            {
                _cachedIcon32 = bmp;
            }

            return bmp;
        }

        private static GraphicsPath CreateRoundedRectanglePath(float x, float y, float width, float height, float radius)
        {
            var path = new GraphicsPath();
            float d = radius * 2;
            path.AddArc(x, y, d, d, 180, 90);
            path.AddArc(x + width - d, y, d, d, 270, 90);
            path.AddArc(x + width - d, y + height - d, d, d, 0, 90);
            path.AddArc(x, y + height - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }
}
