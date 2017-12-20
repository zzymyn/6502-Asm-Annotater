using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace _6502Annotate
{
    public class Annotator
    {
        private static Regex s_IsSprite = new Regex(@"^(sprite|[░█]+(\r?\n[░█]+)*)$", RegexOptions.Multiline | RegexOptions.IgnorePatternWhitespace);

        private ExcelWorksheet m_Worksheet;
        private Dictionary<string, int> m_HeaderCols = new Dictionary<string, int>();
        private Dictionary<string, int> m_LabelRows = new Dictionary<string, int>();

        public static void Annotate(ExcelWorksheet worksheet)
        {
            new Annotator(worksheet);
        }

        private Annotator(ExcelWorksheet worksheet)
        {
            m_Worksheet = worksheet;
            Console.Write("Reading xmlx... ");
            Init();
            Console.WriteLine("done");
            Console.Write("Decode ASM... ");
            DecodeAsms();
            Console.WriteLine("done");
            Console.Write("Link labels... ");
            LinkLabels();
            Console.WriteLine("done");
            Console.Write("Drawing sprites... ");
            DrawSprites();
            Console.WriteLine("done");
            Console.Write("Drawing jump graph... ");
            DrawGraph();
            Console.WriteLine("done");
        }

        private void Init()
        {
            for (int col = 1; col <= m_Worksheet.Dimension.Columns; ++col)
            {
                var header = m_Worksheet.Cells[1, col].Text;
                if (!string.IsNullOrEmpty(header))
                    m_HeaderCols.Add(header, col);
            }

            var labelCol = m_HeaderCols["Label"];

            for (int row = 2; row <= m_Worksheet.Dimension.Rows; ++row)
            {
                var label = m_Worksheet.Cells[row, labelCol].Text;
                if (!string.IsNullOrEmpty(label))
                    m_LabelRows.Add(label, row);
            }
        }

        private void DecodeAsms()
        {
            var asmCol = m_HeaderCols["ASM"];
            var commentCol = m_HeaderCols["Comments"];

            for (int row = 2; row <= m_Worksheet.Dimension.Rows; ++row)
            {
                var asmCell = m_Worksheet.Cells[row, asmCol];
                var commentCell = m_Worksheet.Cells[row, commentCol];

                var decode = AsmUtils.DecodeAsm(asmCell.Text);
                if (string.IsNullOrEmpty(decode))
                    continue;
                commentCell.Value = decode;
            }
        }

        private void LinkLabels()
        {
            var labelCol = m_HeaderCols["Label"];
            var asmCol = m_HeaderCols["ASM"];

            for (int row = 2; row <= m_Worksheet.Dimension.Rows; ++row)
            {
                var asmCell = m_Worksheet.Cells[row, asmCol];

                var label = AsmUtils.GetJumpToLabel(asmCell.Text);
                if (string.IsNullOrEmpty(label))
                    continue;
                int labelRow;
                if (m_LabelRows.TryGetValue(label + ":", out labelRow))
                {
                    var linkCell = m_Worksheet.Cells[labelRow, asmCol];
                    asmCell.Hyperlink = new ExcelHyperLink(linkCell.Address, asmCell.Text);
                    asmCell.StyleName = "Hyperlink";
                }
                else
                {
                    asmCell.Hyperlink = null;
                    asmCell.StyleName = "Normal";
                }
            }
        }

        private void DrawSprites()
        {
            var asmCol = m_HeaderCols["ASM"];
            var commentCol = m_HeaderCols["Comments"];

            for (int row = m_Worksheet.Dimension.Rows; row >= 2; --row)
            {
                var asmCell = m_Worksheet.Cells[row, asmCol];
                var commentCell = m_Worksheet.Cells[row, commentCol];

                if (!s_IsSprite.IsMatch(commentCell.Text))
                    continue;

                // break up bytes to multiple rows:
                var bytes = AsmUtils.GetBytes(asmCell.Text).ToList();
                if (bytes.Count > 1)
                {
                    for (int i = bytes.Count - 1; i >= 1; --i)
                    {
                        m_Worksheet.InsertRow(row + 1, 1);
                        var newAsmCell = m_Worksheet.Cells[row + 1, asmCol];
                        var newCommentCell = m_Worksheet.Cells[row + 1, commentCol];
                        newAsmCell.Value = $".byte ${bytes[i]}";
                        newCommentCell.Value = AsmUtils.DecodeSprite(newAsmCell.Text);
                    }
                }
                asmCell.Value = $".byte ${bytes[0]}";
                commentCell.Value = AsmUtils.DecodeSprite(asmCell.Text); ;
            }
        }

        private void DrawGraph()
        {
            List<HashSet<JumpSet>> graph = CreateGraph();

            int gCol = m_HeaderCols["Tree"];
            for (int row = 2; row <= m_Worksheet.Dimension.Rows; ++row)
            {
                var cell = m_Worksheet.Cells[row, gCol];
                cell.IsRichText = true;
                cell.Value = "";
                var rtc = cell.RichText;

                Color col = Color.Black;
                bool isIn = false;
                bool isOut = false;

                for (int g = graph.Count - 1; g >= 0; --g)
                {
                    var js = graph[g].FirstOrDefault(a => a.Contains(row));

                    if (g <= 5)
                    {
                        if (isOut || isIn)
                        {
                            rtc.Add("─").Color = col;
                        }
                        else
                        {
                            rtc.Add("\u00A0");
                        }
                    }

                    if (js == null)
                    {
                        if (isOut || isIn)
                        {
                            rtc.Add("─").Color = col;
                        }
                        else
                        {
                            rtc.Add("\u00A0");
                        }
                    }
                    else
                    {
                        if (row == js.Min)
                        {
                            col = GraphColor(g);
                            rtc.Add("┌").Color = col;
                        }
                        else if (row == js.Max)
                        {
                            col = GraphColor(g);
                            rtc.Add("└").Color = col;
                        }
                        else if (row == js.Target || js.Sources.Contains(row))
                        {
                            col = GraphColor(g);
                            rtc.Add("├").Color = col;
                        }
                        else
                        {
                            var r = rtc.Add("│");
                            r.Color = GraphColor(g);
                        }

                        if (row == js.Target)
                        {
                            isOut = true;
                        }
                        else if (js.Sources.Contains(row))
                        {
                            isIn = true;
                        }

                    }
                }

                if (isOut)
                {
                    rtc.Add(">").Color = col;
                }
                else if (isIn)
                {
                    rtc.Add("─").Color = col;
                }
                else
                {
                    rtc.Add("\u00A0");
                }
            }
        }

        private List<HashSet<JumpSet>> CreateGraph()
        {
            var graph = new List<HashSet<JumpSet>>();
            var jumpSets = FindJumpSets();

            while (jumpSets.Count > 0)
            {
                var currLevel = new HashSet<JumpSet>();
                foreach (var jumpSet in jumpSets)
                {
                    AddJumpSetToLevel(jumpSet, currLevel);
                }
                graph.Add(currLevel);
                jumpSets.ExceptWith(currLevel);
            }

            return graph;
        }

        private Color GraphColor(int g)
        {
            switch (g)
            {
                case 0:
                    return Color.Red;
                case 1:
                    return Color.Blue;
                case 2:
                    return Color.Green;
                case 3:
                    return Color.Orange;
                case 4:
                    return Color.Magenta;
                case 5:
                    return Color.DarkCyan;
                default:
                    return Color.Gray;
            }
        }

        private Color GraphColorJoin(Color col1, Color col2)
        {
            var r = (int)Math.Floor(0.5f * col1.R + 0.5f * col2.R);
            var g = (int)Math.Floor(0.5f * col1.G + 0.5f * col2.G);
            var b = (int)Math.Floor(0.5f * col1.B + 0.5f * col2.B);
            return Color.FromArgb(255, r, g, b);
        }

        private void AddJumpSetToLevel(JumpSet js, HashSet<JumpSet> level)
        {
            var rejected = new HashSet<JumpSet>();

            foreach (var existing in level)
            {
                if (!existing.Overlaps(js))
                    continue;

                if (existing.IsSupersetOf(js) || existing.Sources.Contains(js.Target) || js.Min < existing.Min && js.Max < existing.Max)
                {
                    rejected.Add(existing);
                }
                else
                {
                    return;
                }
            }

            level.ExceptWith(rejected);
            level.Add(js);
        }

        private HashSet<JumpSet> FindJumpSets()
        {
            var jumpSets = new Dictionary<int, JumpSet>();
            var labelCol = m_HeaderCols["Label"];
            var asmCol = m_HeaderCols["ASM"];

            for (int row = 2; row <= m_Worksheet.Dimension.Rows; ++row)
            {
                var asmCell = m_Worksheet.Cells[row, asmCol];
                var label = AsmUtils.GetJumpToLabel(asmCell.Text);
                if (string.IsNullOrEmpty(label))
                    continue;
                int labelRow;
                if (!m_LabelRows.TryGetValue(label + ":", out labelRow))
                    continue;

                JumpSet jumpSet;
                if (!jumpSets.TryGetValue(labelRow, out jumpSet))
                {
                    jumpSet = new JumpSet(labelRow);
                    jumpSets[labelRow] = jumpSet;
                }
                jumpSet.AddSource(row);
            }

            return new HashSet<JumpSet>(jumpSets.Values);
        }
    }
}
