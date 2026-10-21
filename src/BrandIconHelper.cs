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

        private static readonly System.Collections.Generic.Dictionary<string, Bitmap> _cachedSmallIcons =
            new System.Collections.Generic.Dictionary<string, Bitmap>(StringComparer.OrdinalIgnoreCase);
        private static readonly object _iconLock = new object();

        public static Bitmap GetSmallIcon(string controlId)
        {
            if (string.IsNullOrEmpty(controlId)) controlId = "default";
            lock (_iconLock)
            {
                Bitmap cached;
                if (_cachedSmallIcons.TryGetValue(controlId, out cached) && cached != null)
                {
                    return cached;
                }

                var bmp = RenderSmallIcon(controlId);
                _cachedSmallIcons[controlId] = bmp;
                return bmp;
            }
        }

        private static Bitmap RenderSmallIcon(string id)
        {
            int size = 16;
            var bmp = new Bitmap(size, size, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);

                // 1. 底色渐变圆角底 (#107C41 -> #094B26)
                using (var path = CreateRoundedRectanglePath(0.75f, 0.75f, 14.5f, 14.5f, 2.5f))
                {
                    using (var brush = new LinearGradientBrush(
                        new PointF(0, 0),
                        new PointF(0, size),
                        ColorTranslator.FromHtml("#107C41"),
                        ColorTranslator.FromHtml("#094B26")))
                    {
                        g.FillPath(brush, path);
                    }

                    using (var pen = new Pen(ColorTranslator.FromHtml("#2EB86F"), 0.8f))
                    {
                        g.DrawPath(pen, path);
                    }
                }

                // 2. 白色线性符号
                using (var whitePen = new Pen(Color.White, 1.2f))
                using (var whiteBrush = new SolidBrush(Color.White))
                {
                    whitePen.StartCap = LineCap.Round;
                    whitePen.EndCap = LineCap.Round;
                    whitePen.LineJoin = LineJoin.Round;

                    string key = id.ToLowerInvariant();
                    switch (key)
                    {
                        case "btnmacrolibrary":
                        case "fav_lib":
                        case "fav_open_lib":
                            // 代码块 </>
                            g.DrawLines(whitePen, new PointF[] { new PointF(5.2f, 5.0f), new PointF(3.2f, 8.0f), new PointF(5.2f, 11.0f) });
                            g.DrawLine(whitePen, 6.8f, 11.5f, 9.2f, 4.5f);
                            g.DrawLines(whitePen, new PointF[] { new PointF(10.8f, 5.0f), new PointF(12.8f, 8.0f), new PointF(10.8f, 11.0f) });
                            break;

                        case "menufavoritemacros":
                        case "favorite":
                            // 五角星
                            PointF[] starPts = CalculateStarPoints(8.0f, 8.0f, 4.6f, 1.9f);
                            g.FillPolygon(whiteBrush, starPts);
                            break;

                        case "btndatatools":
                            // 扳手
                            using (var handlePen = new Pen(Color.White, 1.8f))
                            {
                                handlePen.StartCap = LineCap.Round;
                                g.DrawLine(handlePen, 4.2f, 11.8f, 9.5f, 6.5f);
                            }
                            g.DrawArc(whitePen, 8.0f, 3.5f, 4.5f, 4.5f, 220, 260);
                            break;

                        case "btnbatchmodal":
                            // 三张叠放的矩形
                            g.DrawRectangle(whitePen, 3.0f, 3.0f, 6.0f, 4.5f);
                            g.DrawRectangle(whitePen, 5.0f, 5.5f, 6.0f, 4.5f);
                            using (var fillBrush = new SolidBrush(ColorTranslator.FromHtml("#107C41")))
                            {
                                g.FillRectangle(fillBrush, 7.0f, 8.0f, 6.0f, 4.5f);
                            }
                            g.DrawRectangle(whitePen, 7.0f, 8.0f, 6.0f, 4.5f);
                            break;

                        case "btnworkflowmodal":
                            // 两个节点 + 一条连线
                            g.FillEllipse(whiteBrush, 3.0f, 6.5f, 3.2f, 3.2f);
                            g.FillEllipse(whiteBrush, 9.8f, 6.5f, 3.2f, 3.2f);
                            g.DrawLine(whitePen, 6.0f, 8.1f, 10.0f, 8.1f);
                            break;

                        case "btnapisettings":
                            // 齿轮
                            g.DrawEllipse(whitePen, 5.0f, 5.0f, 6.0f, 6.0f);
                            g.FillEllipse(whiteBrush, 7.0f, 7.0f, 2.0f, 2.0f);
                            g.DrawLine(whitePen, 8.0f, 3.2f, 8.0f, 5.0f);
                            g.DrawLine(whitePen, 8.0f, 11.0f, 8.0f, 12.8f);
                            g.DrawLine(whitePen, 3.2f, 8.0f, 5.0f, 8.0f);
                            g.DrawLine(whitePen, 11.0f, 8.0f, 12.8f, 8.0f);
                            break;

                        case "fav_run":
                            // 运行三角形
                            PointF[] playPts = new PointF[]
                            {
                                new PointF(5.5f, 4.5f),
                                new PointF(11.5f, 8.0f),
                                new PointF(5.5f, 11.5f)
                            };
                            g.FillPolygon(whiteBrush, playPts);
                            break;

                        case "fav_empty":
                            // 提示信息 i
                            g.DrawEllipse(whitePen, 3.5f, 3.5f, 9.0f, 9.0f);
                            g.FillEllipse(whiteBrush, 7.2f, 5.5f, 1.6f, 1.6f);
                            g.DrawLine(whitePen, 8.0f, 8.0f, 8.0f, 10.5f);
                            break;

                        default:
                            g.FillEllipse(whiteBrush, 6.0f, 6.0f, 4.0f, 4.0f);
                            break;
                    }
                }
            }
            return bmp;
        }

        private static PointF[] CalculateStarPoints(float cx, float cy, float rOut, float rIn)
        {
            PointF[] pts = new PointF[10];
            double angleStep = Math.PI / 5.0;
            double startAngle = -Math.PI / 2.0;
            for (int i = 0; i < 10; i++)
            {
                float r = (i % 2 == 0) ? rOut : rIn;
                double a = startAngle + i * angleStep;
                pts[i] = new PointF(cx + (float)(r * Math.Cos(a)), cy + (float)(r * Math.Sin(a)));
            }
            return pts;
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
