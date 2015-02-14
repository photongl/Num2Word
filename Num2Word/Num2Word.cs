/********************************************************************
 Copyright (c) 2015. Akshay Mishra. All rights reserved.

Licensed under the Apache License, Version 2.0 (the "License"); you may not use
these files except in compliance with the License. You may obtain a copy of the
License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed
under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR
CONDITIONS OF ANY KIND, either express or implied. See the License for the
specific language governing permissions and limitations under the License.
********************************************************************/

using System;
using System.Collections.Generic;
using System.Text;
using ExcelDna.Integration;

namespace Num2Word
{
    public static class Num2Word
    {
        static string[] ONES = { "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE" };
        static string[] TENS = { "TEN", "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY", "SEVENTY", "EIGHTY", "NINETY" };
        static string[] TEENS = { "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN", "SIXTEEN", "SEVENTEEN", "EIGHTEEN", "NINETEEN" };
        static string HUNDRED = "HUNDRED";
        static string THOUSAND = "THOUSAND";
        static string LAKH = "LAKH";
        static string CRORE = "CRORE";

        [ExcelFunction(Description="Converts a number to word using the Hindi numeral system")]
        public static string N2W(Int64 num)
        {
            StringBuilder sb = new StringBuilder();

            Int64 place_crore = num / 10000000;
            Int64 rem_crore = num % 10000000;
            Int64 place_lakh = rem_crore / 100000;
            Int64 rem_lakh = rem_crore % 100000;
            Int64 place_thousand = rem_lakh / 1000;
            Int64 rem_thousand = rem_lakh % 1000;
            Int64 place_hundred = rem_thousand / 100;
            Int64 rem_hundred = rem_thousand % 100;
            Int64 place_ten = rem_hundred / 10;
            Int64 rem_ten = rem_hundred % 10;

            if (place_crore != 0) sb.AppendFormat("{0} {1} ", N2W(place_crore), CRORE);
            if (place_lakh != 0) sb.AppendFormat("{0} {1} ", N2W(place_lakh), LAKH);
            if (place_thousand != 0) sb.AppendFormat("{0} {1} ", N2W(place_thousand), THOUSAND);
            if (place_hundred != 0) sb.AppendFormat("{0} {1} ", ONES[place_hundred - 1]
                            , (rem_hundred == 0) ? HUNDRED : String.Format("{0} AND ", HUNDRED));

            // handle 11-19
            if (place_ten == 1 && rem_ten != 0)
            {
                sb.AppendFormat("{0}", TEENS[rem_ten - 1]);
            }
            else
            {
                if (place_ten != 0) sb.AppendFormat("{0} ", TENS[place_ten - 1]);
                if (rem_ten != 0) sb.AppendFormat("{0}", ONES[rem_ten - 1]);
            }

            return sb.ToString();
        }
    }
}
