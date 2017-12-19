using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace _6502Annotate
{
    public class Annotator
    {
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
            Init();
            DecodeAsms();
            LinkLabels();
            DrawGraph();
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
                    var labelCell = m_Worksheet.Cells[labelRow, labelCol];
                    asmCell.Hyperlink = new ExcelHyperLink(labelCell.Address, asmCell.Text);
                    asmCell.StyleName = "Hyperlink";
                }
                else
                {
                    asmCell.Hyperlink = null;
                    asmCell.StyleName = "Normal";
                }
            }
        }

        private void DrawGraph()
        {
            var gens = new List<HashSet<JumpSet>>();
            var jumpSets = FindJumpSets();

            while (jumpSets.Count > 0)
            {
                var currGen = new HashSet<JumpSet>();
                foreach (var jumpSet in jumpSets)
                {
                    AddJumpSet(jumpSet, currGen);
                }
                gens.Add(currGen);
                jumpSets.ExceptWith(currGen);
            }

            int gCol = m_HeaderCols["Tree"];
            for (int row = 2; row <= m_Worksheet.Dimension.Rows; ++row)
            {
                var sb = new StringBuilder();

                bool isIn = false;
                bool isOut = false;

                for (int g = gens.Count - 1; g >= 0; --g)
                {
                    var js = gens[g].FirstOrDefault(a => a.Contains(row));

                    if (isOut || isIn)
                    {
                        sb.Append('─');
                    }
                    else
                    {
                        sb.Append(' ');
                    }

                    if (js == null)
                    {
                        if (isOut || isIn)
                        {
                            sb.Append('─');
                        }
                        else
                        {
                            sb.Append(' ');
                        }
                    }
                    else
                    {
                        if (row == js.Min)
                        {
                            if (isIn || isOut)
                                sb.Append('┬');
                            else
                                sb.Append('┌');
                        }
                        else if (row == js.Max)
                        {
                            if (isIn || isOut)
                                sb.Append('┴');
                            else
                                sb.Append('└');
                        }
                        else if (row == js.Target || js.Sources.Contains(row))
                        {
                            sb.Append('├');
                        }
                        else
                        {
                            sb.Append('│');
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
                    sb.Append('>');
                }
                else if (isIn)
                {
                    sb.Append('─');
                }
                else
                {
                    sb.Append(' ');
                }

                m_Worksheet.Cells[row, gCol].Value = sb.ToString();
            }
        }

        private void AddJumpSet(JumpSet js, HashSet<JumpSet> gen)
        {
            var rejected = new HashSet<JumpSet>();

            foreach (var existing in gen)
            {
                if (!existing.Overlaps(js))
                    continue;

                if (existing.IsSupersetOf(js) || js.Min < existing.Min && js.Max < existing.Max)
                {
                    rejected.Add(existing);
                }
                else
                {
                    return;
                }
            }

            gen.ExceptWith(rejected);
            gen.Add(js);
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
