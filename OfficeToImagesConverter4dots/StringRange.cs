using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeToImagesConverter4dots
{
    public class StringRange
    {
        private string Range = "";

        public StringRange(string stringrange)
        {
            Range = stringrange;
        }

        public bool IsInRange(int k)
        {
            if (Range == string.Empty) return true;           

            string[] ranges = Range.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

            string kk = k.ToString();

            for (int m = 0; m < ranges.Length; m++)
            {
                // if only a range for all sheets e.g. A1:C10
                try
                {
                    int sq1pos0 = ranges[m].IndexOf(":");
                    int sq2pos0 = ranges[m].IndexOf(":", sq1pos0 + 1);

                    if (sq1pos0 >= 0 && sq2pos0 < 0)
                    {
                        return true;
                    }
                }
                catch { }

                // sheet number

                if (ranges[m] == kk) return true;

                if (ranges[m].IndexOf("-") > 0)
                {
                    string st = ranges[m].Substring(0, ranges[m].IndexOf("-"));
                    int ist = -1;

                    ist = int.Parse(st);

                    string en = ranges[m].Substring(ranges[m].IndexOf("-") + 1);
                    int ien = -1;

                    ien = int.Parse(en);

                    if (k >= ist && k <= ien)
                    {
                        return true;
                    }
                }
                else
                {
                    // sheet number with sheet range
                    try
                    {
                        int sq1pos = ranges[m].IndexOf(":");
                        int sq2pos = ranges[m].IndexOf(":", sq1pos + 1);

                        if (sq1pos >= 0 && sq2pos >= 0)
                        {
                            if (ranges[m].StartsWith(kk + ":"))
                            {
                                return true;
                            }                            
                        }
                    }
                    catch { }                    
                }
            }

            return false;
        }
    }
}
