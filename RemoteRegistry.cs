using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace TechTool
{
    class RemoteRegistry
    {
        public String[] regQuery(String regLocation)
        {
            String tempResult = "";
            String[] tempArray;
            PowerShell ps = PowerShell.Create();

            String tempString = ReplaceNonPrintableCharacters(regLocation, "");
            //System.Windows.Forms.MessageBox.Show("RegQuery : " + "reg query " + tempString);
            ps.AddScript("reg query " + tempString);
            Collection<PSObject> results = ps.Invoke();

            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject psObject in results)
            {
                stringBuilder.AppendLine(psObject.ToString());
            }
            tempResult = stringBuilder.ToString();

            tempArray = tempResult.Split('\n');
            return tempArray;
        }

        public String[] regQueryValue(String regLocation)
        {
            String tempResult = "";
            String[] tempArray;
            PowerShell ps = PowerShell.Create();

            String tempString = ReplaceNonPrintableCharacters(regLocation, "");
            //System.Windows.Forms.MessageBox.Show("RegQueryValue : " + "reg query " + tempString + " /ve");
            ps.AddScript("reg query " + tempString + " /ve");
            Collection<PSObject> results = ps.Invoke();

            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject psObject in results)
            {
                stringBuilder.AppendLine(psObject.ToString());
            }
            tempResult = stringBuilder.ToString();

            tempArray = tempResult.Split('\n');
            return tempArray;
        }
        public String ReplaceNonPrintableCharacters(string s, string replaceWith)
        {
            StringBuilder result = new StringBuilder();
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                byte b = (byte)c;
                if (b < 32)
                    result.Append(replaceWith);
                else
                    result.Append(c);
            }
            return result.ToString();
        }
    }
}
