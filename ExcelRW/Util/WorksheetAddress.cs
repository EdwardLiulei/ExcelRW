using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelReadAndWrite.Util
{
    public class WorksheetAddress
    {

        public static string GetColumnAddress(int index)
        {
            if (index <= 0)
                throw new Exception(string.Format("Illegal index :{0}",index));
            string address = "";
            int num = index;
            while (num > 26)
            {
                address = ((char)(num % 26 + 64)).ToString() + address;
                num = num / 26;
            }

            address = ((char)(num + 64)).ToString() + address;
            return address;
        }

        public static int GetColumnIndex(string address)
        {
            if (!Regex.IsMatch(address, @"^\w+$"))
                throw new Exception(string.Format("Illegal address: {0}",address));

            int num = 0;
            char[] charArray = address.ToUpper().ToCharArray();
            for (int i = 0; i < charArray.Length; i++)
            {
                int charIndex = (int)charArray[i]-64;
                num += charIndex * (int)Math.Pow(26,charArray.Length-i-1);
            }

            return num;
        }

        
    }
}
