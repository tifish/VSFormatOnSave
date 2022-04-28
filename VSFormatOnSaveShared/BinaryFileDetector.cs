using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Tinyfish.FormatOnSave
{
    static class BinaryFileDetector
    {
        private static readonly HashSet<string> TextFileExtensions;

        static BinaryFileDetector()
        {
            TextFileExtensions = new HashSet<string>(
                @".cs;.vb;.resx;.xsd;.wsdl;.htm;.html;.aspx;.ascx;.asmx;.svc;.asax;.config;.asp;.asa;.cshtml;.vbhtml;.razor;.css;.xml
.c;.cpp;.cxx;.cc;.tli;.tlh;.h;.hh;.hpp;.hxx;.hh;.inl;.ipp;.rc;.resx;.idl;.asm;.inc
.txt;.log;.md
.ini;.yaml;.json".Split(';', '\n'));
        }

        public static bool IsBinary(string path)
        {
            var ext = Path.GetExtension(path).ToLower();
            if (TextFileExtensions.Contains(ext))
                return false;

            var length = new FileInfo(path).Length;
            if (length == 0)
                return true;

            using (var fs = new FileStream(path, FileMode.Open))
            {
                if (IsUnicodeFile(fs))
                    return false;

                return TestBinary(fs);
            }
        }

        private static readonly byte[][] UnicodeBOMs =
        {
            Encoding.UTF8.GetPreamble(),
            Encoding.Unicode.GetPreamble(),
            Encoding.BigEndianUnicode.GetPreamble(),
            Encoding.UTF32.GetPreamble(),
        };

        private static bool IsUnicodeFile(FileStream fs)
        {
            var buf = new byte[16];
            fs.Seek(0, SeekOrigin.Begin);
            fs.Read(buf, 0, 16);

            foreach (var bom in UnicodeBOMs)
                if (HasBom(buf, bom))
                    return true;

            return false;
        }

        private static bool HasBom(byte[] header, byte[] bom)
        {
            for (var i = 0; i < bom.Length; i++)
                if (header[i] != bom[i])
                    return false;

            return true;
        }

        private static bool TestBinary(FileStream fs)
        {
            int ch;
            while ((ch = fs.ReadByte()) != -1)
                if (IsControlChar(ch))
                    return true;

            return false;
        }

        public static bool IsControlChar(int ch)
        {
            return (ch > Chars.NUL && ch < Chars.BS)
                   || (ch > Chars.CR && ch < Chars.SUB);
        }

        public static class Chars
        {
            public static char NUL = (char) 0; // Null char
            public static char BS = (char) 8; // Back Space
            public static char CR = (char) 13; // Carriage Return
            public static char SUB = (char) 26; // Substitute
        }
    }
}
