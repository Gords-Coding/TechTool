using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text;
using System.Threading.Tasks;

namespace TechTool
{
    class AccountManager
    {
        public Boolean accountExist(String domainInput, String usernameInput)
        {
            String tempUsername = ReplaceNonPrintableCharacters(usernameInput, "");

            String psCommand = "Get-ADUser -server " + domainInput + " -LDAPFilter \"(sAMAccountName=" + tempUsername + ")\" | Select Enabled";
            Boolean returnBoolean = true;

            PowerShell ps = PowerShell.Create();
            ps.AddScript(psCommand);
            Collection<PSObject> results = ps.Invoke();

            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject psObject in results)
            {
                stringBuilder.AppendLine(psObject.ToString());
            }

            if (stringBuilder.ToString().Length < 1)
            {
                returnBoolean = false;
            }
            return returnBoolean;
        }
        public String getDisplayName(String domainInput, String usernameInput)
        {
            String returnString = "";
            String tempString = "";
            String tempUsername = ReplaceNonPrintableCharacters(usernameInput, "");

            String psCommand = "Get-ADUser -server " + domainInput + " -LDAPFilter \"(sAMAccountName=" + tempUsername + ")\" -Properties displayName";
            String[] tempArray;
            //System.Windows.Forms.MessageBox.Show(psCommand);

            PowerShell ps = PowerShell.Create();
            ps.AddScript(psCommand);
            ps.AddCommand("Out-String");
            Collection<PSObject> results = ps.Invoke();

            StringBuilder stringBuilder = new StringBuilder();

            foreach (PSObject psObject in results)
            {
                stringBuilder.AppendLine(psObject.ToString());
            }
            tempString = stringBuilder.ToString();
            tempArray = tempString.Split('\r');
            //System.Windows.Forms.MessageBox.Show("tempArray[2] : " + tempArray[2]);
            returnString = tempArray[2];
            return returnString;
        }

        public String getPasswordLastSet(String domainInput, String usernameInput)
        {
            String returnString = "";
            //String tempString = "";
            String tempUsername = ReplaceNonPrintableCharacters(usernameInput, "");
            String psCommand = "Get-ADUser -Identity " + usernameInput + " -Server " + domainInput + " -Properties PasswordLastSet | select PasswordLastSet";

            String[] tempArray;

            PowerShell ps = PowerShell.Create();
            ps.AddScript(psCommand);
            ps.AddCommand("Out-String");
            Collection<PSObject> results = ps.Invoke();

            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject psObject in results)
            {
                stringBuilder.AppendLine(psObject.ToString());
            }

            //System.Windows.Forms.MessageBox.Show(stringBuilder.ToString());

            tempArray = stringBuilder.ToString().Split('\r');
            //returnString = stringBuilder.ToString();
            returnString = tempArray[3];
            return returnString;
        }
        public String getPasswordExpiringDays(String domainInput, String usernameInput)
        {
            String returnString = "";
            String tempUsername = ReplaceNonPrintableCharacters(usernameInput, "");
            String psCommand = "(([datetime]::FromFileTime((Get-ADUser -Identity " + tempUsername + " -Server " + domainInput + " -Properties \"msDS-UserPasswordExpiryTimeComputed\").\"msDS-UserPasswordExpiryTimeComputed\"))-(Get-Date)).Days";

            //System.Windows.Forms.MessageBox.Show("Domain : " + domainInput + " username : " + usernameInput);

            PowerShell ps = PowerShell.Create();
            ps.AddScript(psCommand);
            Collection<PSObject> results = ps.Invoke();

            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject psObject in results)
            {
                stringBuilder.AppendLine(psObject.ToString());
            }

            //System.Windows.Forms.MessageBox.Show(stringBuilder.ToString());
            returnString = stringBuilder.ToString();

            return returnString;
        }

        public String getEmployeeID(String domainInput, String usernameInput)
        {
            String returnString = "";
            String tempUsername = ReplaceNonPrintableCharacters(usernameInput, "");



            String psCommand = "Get-ADUser -LDAPFilter \"(sAMAccountName=" + tempUsername + ")\" -Properties EmployeeID -Server " + domainInput + " | select employeeID";
            String[] tempArray;

            String fbuUsername = "";

            PowerShell ps = PowerShell.Create();
            ps.AddScript(psCommand);
            ps.AddCommand("Out-String");
            Collection<PSObject> results = ps.Invoke();

            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject psObject in results)
            {
                stringBuilder.AppendLine(psObject.ToString());
            }
            //System.Windows.Forms.MessageBox.Show("StringBuilder.ToString : " + stringBuilder.ToString());
            tempArray = stringBuilder.ToString().Split('\r');
            
            fbuUsername = tempArray[3].Replace(" ", "");
            //fbuUsername = tempArray2[1];
            returnString = fbuUsername.Replace("\n", "");
            //returnString = stringBuilder.ToString();

            return returnString;
        }
        
        public Boolean getAccountEnabledStatus(String domainInput, String usernameInput)
        {
            Boolean returnBoolean = false;
            String tempUsername = ReplaceNonPrintableCharacters(usernameInput, "");
            String tempString = "";

            String psCommand = "Get-ADUser -server " + domainInput + " -LDAPFilter \"(sAMAccountName=" + tempUsername + ")\" | Select Enabled";

            PowerShell ps = PowerShell.Create();
            ps.AddScript(psCommand);
            Collection<PSObject> results = ps.Invoke();


            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject psObject in results)
            {
                stringBuilder.AppendLine(psObject.ToString());
            }

            tempString = stringBuilder.ToString();
            if (tempString.Contains("True"))
            {
                returnBoolean = true;
            }else if (tempString.Contains("False"))
            {
                returnBoolean = false;
            }
            return returnBoolean;
        }
        public Boolean getPasswordNeverExpiresStatus(String domainInput, String usernameInput)
        {
            Boolean returnBoolean = false;
            String tempUsername = ReplaceNonPrintableCharacters(usernameInput, "");
            String tempString = "";

            String psCommand = "Get-ADUser -Identity " + tempUsername + " -Server " + domainInput + " -properties PasswordNeverExpires | Select-Object PasswordNeverExpires";

            PowerShell ps = PowerShell.Create();
            ps.AddScript(psCommand);
            Collection<PSObject> results = ps.Invoke();

            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject psObject in results)
            {
                stringBuilder.AppendLine(psObject.ToString());
            }

            tempString = stringBuilder.ToString();

            if (tempString.Contains("True"))
            {
                returnBoolean = true;
            }
            else if (tempString.Contains("False"))
            {
                returnBoolean = false;
            }
            return returnBoolean;
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
