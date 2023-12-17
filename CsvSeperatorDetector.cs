using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMergeSelectedRows
{
    static class CsvSeperatorDetector
    {
        private static readonly char[] SeparatorChars = { ';', '|', '\t', ',' };

        public static char DetectSeparator(string csvFilePath)
        {
            string[] lines = File.ReadAllLines(csvFilePath);
            return DetectSeparator(lines);
        }

        public static char DetectSeparator(string[] lines)
        {
            var q = SeparatorChars.Select(
                sep => new{ 
                xSeparator = sep, xFound = lines.GroupBy(line => line
                    .Count(ch => ch == sep)
                        ) 
                }
            )
                .OrderByDescending(res => res.xFound.Count(grp => grp.Key > 0))
                .ThenBy(res => res.xFound.Count())
                .First();

            return q.xSeparator;
        }
    }
}
