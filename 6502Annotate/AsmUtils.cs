using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace _6502Annotate
{
    public static class AsmUtils
    {
        private static ValueTuple<Regex, string>[] s_ArgDecoders =
        {
            ValueTuple.Create(new Regex(@"^(\w+)$"), @"$1"),
            ValueTuple.Create(new Regex(@"^\$([0-9A-F]+)$"), @"mem[$$$1]"),
            ValueTuple.Create(new Regex(@"^\#\$([0-9A-F]+)$"), @"$$$1"),
            ValueTuple.Create(new Regex(@"^(\w+),([XY])$"), @"$1[$2]"),
            ValueTuple.Create(new Regex(@"^\$([0-9A-F]+),([XY])$"), @"mem[$$$1 + $2]"),
            ValueTuple.Create(new Regex(@"^\(\$([0-9A-F]+)\),([XY])$"), @"mem[$$$1][$2]"),
        };

        private static ValueTuple<Regex, string>[] s_AsmDecoders =
        {
            ValueTuple.Create(new Regex(@"^SEI$"), @"set interrupt disable"),
            ValueTuple.Create(new Regex(@"^CLD$"), @"clear decimal mode"),
            ValueTuple.Create(new Regex(@"^SED$"), @"set decimal mode"),
            ValueTuple.Create(new Regex(@"^SEC$"), @"carry = 1"),
            ValueTuple.Create(new Regex(@"^CLC$"), @"carry = 0"),
            ValueTuple.Create(new Regex(@"^RTS$"), @"return"),
            ValueTuple.Create(new Regex(@"^NOP$"), @""),
            ValueTuple.Create(new Regex(@"^DE([XY])$"), @"--{1}"),
            ValueTuple.Create(new Regex(@"^IN([XY])$"), @"++{1}"),
            ValueTuple.Create(new Regex(@"^INC (.*)$"), @"++{1}"),
            ValueTuple.Create(new Regex(@"^DEC (.*)$"), @"--{1}"),
            ValueTuple.Create(new Regex(@"^BEQ (.*)$"), @"if (prev == 0) goto {1}"),
            ValueTuple.Create(new Regex(@"^BNE (.*)$"), @"if (prev != 0) goto {1}"),
            ValueTuple.Create(new Regex(@"^BMI (.*)$"), @"if (prev < 0) goto {1}"),
            ValueTuple.Create(new Regex(@"^BPL (.*)$"), @"if (prev >= 0) goto {1}"),
            ValueTuple.Create(new Regex(@"^BCC (.*)$"), @"if (carry == 0) goto {1}"),
            ValueTuple.Create(new Regex(@"^BCS (.*)$"), @"if (carry != 0) goto {1}"),
            ValueTuple.Create(new Regex(@"^JMP (.*)$"), @"goto {1}"),
            ValueTuple.Create(new Regex(@"^JSR (.*)$"), @"call {1}"),
            ValueTuple.Create(new Regex(@"^AND (.*)$"), @"A &= {1}"),
            ValueTuple.Create(new Regex(@"^ORA (.*)$"), @"A |= {1}"),
            ValueTuple.Create(new Regex(@"^EOR (.*)$"), @"A ^= {1}"),
            ValueTuple.Create(new Regex(@"^ADC (.*)$"), @"A += {1} + carry"),
            ValueTuple.Create(new Regex(@"^SBC (.*)$"), @"A -= {1} + ~carry"),
            ValueTuple.Create(new Regex(@"^CMP (.*)$"), @"cmp(A, {1})"),
            ValueTuple.Create(new Regex(@"^ROL (.*)$"), @"({1} <<= 1) |= carry"),
            ValueTuple.Create(new Regex(@"^ROR (.*)$"), @"({1} >>= 1) |= (carry << 7)"),
            ValueTuple.Create(new Regex(@"^CP([XYA]) (.*)$"), @"cmp({1}, {2})"),
            ValueTuple.Create(new Regex(@"^T([XYAS])([XYAS])$"), @"{2} = {1}"),
            ValueTuple.Create(new Regex(@"^LD([XYA]) (.*)$"), @"{1} = {2}"),
            ValueTuple.Create(new Regex(@"^ASL (.*)$"), @"{1} <<= 1"),
            ValueTuple.Create(new Regex(@"^LSR (.*)$"), @"{1} >>= 1"),
            ValueTuple.Create(new Regex(@"^ST([XYA]) (.*)$"), @"{2} = {1}"),
        };

        private static Regex[] s_JumpAsms =
        {
            new Regex(@"^BEQ (.*)$"),
            new Regex(@"^BNE (.*)$"),
            new Regex(@"^BMI (.*)$"),
            new Regex(@"^BPL (.*)$"),
            new Regex(@"^BCC (.*)$"),
            new Regex(@"^BCS (.*)$"),
            new Regex(@"^JMP (.*)$"),
            new Regex(@"^JSR (.*)$"),
        };

        public static string DecodeAsm(string asm)
        {
            foreach (var asmDecoder in s_AsmDecoders)
            {
                var m = asmDecoder.Item1.Match(asm);
                if (!m.Success)
                    continue;
                return string.Format(asmDecoder.Item2, m.Groups.OfType<Group>().Select(a => DecodeArg(a.Value)).ToArray());
            }
            return null;
        }

        public static string DecodeArg(string arg)
        {
            foreach (var argDecoder in s_ArgDecoders)
            {
                if (argDecoder.Item1.IsMatch(arg))
                {
                    return argDecoder.Item1.Replace(arg, argDecoder.Item2);
                }
            }
            return "err";
        }

        public static string GetJumpToLabel(string asm)
        {
            foreach (var jumpAsm in s_JumpAsms)
            {
                var m = jumpAsm.Match(asm);
                if (!m.Success)
                    continue;
                return m.Groups[1].Value;
            }
            return null;
        }
    }
}
