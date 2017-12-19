using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _6502Annotate
{
    internal class JumpSet
    {
        public int Target { get; set; }
        public HashSet<int> Sources { get; set; } = new HashSet<int>();

        public int Min
        {
            get { return Math.Min(Target, Sources.Min()); }
        }

        public int Max
        {
            get { return Math.Max(Target, Sources.Max()); }
        }

        public JumpSet(int target)
        {
            Target = target;
        }

        public void AddSource(int source)
        {
            Sources.Add(source);
        }

        public bool IsSupersetOf(JumpSet o)
        {
            return Min <= o.Min && Max >= o.Max;
        }

        public bool Overlaps(JumpSet o)
        {
            return Max >= o.Min && Min <= o.Max;
        }

        public bool Contains(int row)
        {
            return row >= Min && row <= Max;
        }
    }
}
