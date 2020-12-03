using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TechTool
{
    class FBITHelpTool
    {
        public String getComputerModel(String input)
        {
            String tempResult = "";
            PowerShell ps = PowerShell.Create();
            ps.AddScript("Invoke-Command -ComputerName " + input + " {Get-ItemPropertyValue -Path HKLM:\\HARDWARE\\DESCRIPTION\\System\\BIOS -Name SystemVersion}");
            Collection<PSObject> results = ps.Invoke();

            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject psObject in results)
            {
                stringBuilder.AppendLine(psObject.ToString());
            }

            tempResult = stringBuilder.ToString();
            return tempResult;
        }

        public String getComputerMake(String input)
        {
            String tempResult = "";
            PowerShell ps = PowerShell.Create();
            ps.AddScript("Invoke-Command -ComputerName " + input + " {Get-ItemPropertyValue -Path HKLM:\\HARDWARE\\DESCRIPTION\\System\\BIOS -Name SystemManufacturer}");
            Collection<PSObject> results = ps.Invoke();

            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject psObject in results)
            {
                stringBuilder.AppendLine(psObject.ToString());
            }

            tempResult = stringBuilder.ToString();
            return tempResult;
        }

        public String getComputerModelType(String input)
        {
            String tempResult = "";
            PowerShell ps = PowerShell.Create();
            ps.AddScript("Invoke-Command -ComputerName " + input + " {Get-ItemPropertyValue -Path HKLM:\\HARDWARE\\DESCRIPTION\\System\\BIOS -Name BaseBoardProduct}");
            Collection<PSObject> results = ps.Invoke();

            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject psObject in results)
            {
                stringBuilder.AppendLine(psObject.ToString());
            }

            tempResult = stringBuilder.ToString();
            return tempResult;
        }

        public String getComputerOS(String input)
        {
            String tempResult = "";
            PowerShell ps = PowerShell.Create();
            ps.AddScript("Invoke-Command -ComputerName " + input + " {Get-ItemPropertyValue -Path \'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\' -Name ProductName}");
            Collection<PSObject> results = ps.Invoke();

            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject psObject in results)
            {
                stringBuilder.AppendLine(psObject.ToString());
            }

            tempResult = stringBuilder.ToString();
            return tempResult;
        }

        public String getReleaseName(String input)
        {
            String tempString = "";
            String psCommand = "Invoke-Command -ComputerName " + input + " {Get-ItemPropertyValue -Path \'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\' -Name ReleaseId}";

            PowerShell ps = PowerShell.Create();
            ps.AddScript(psCommand);
            Collection<PSObject> results = ps.Invoke();

            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject psObject in results)
            {
                stringBuilder.AppendLine(psObject.ToString());
            }

            tempString = stringBuilder.ToString();
            return tempString;
        }

        public String GetRestartPending(String asset)
        {
            Cursor.Current = Cursors.WaitCursor;
            //Boolean tempBoolean = false;
            String psScriptLocation = @"AdditionalFiles\Test-PendingReboot.ps1";
            //String psCommand = "Reg Query \"\\\\" + input + "\\HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\" /v ReleaseId";
            String tempString = "";
            
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();

            Pipeline pipeline = runspace.CreatePipeline();
            Command command = new Command(psScriptLocation);
            command.Parameters.Add(null, asset);
            pipeline.Commands.Add(command);

            Collection<PSObject> results = pipeline.Invoke();
            runspace.Close();

            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject psObject in results)
            {
                stringBuilder.AppendLine(psObject.ToString());
            }

            //System.Windows.Forms.MessageBox.Show(stringBuilder.ToString());
            if (stringBuilder.ToString().Contains("True"))
            {
                //System.Windows.Forms.MessageBox.Show("Contains True");
                tempString = "Reboot Required";
            }
            else if(stringBuilder.ToString().Contains("False"))
            {
                //System.Windows.Forms.MessageBox.Show("Contains False");
                tempString = "No Reboot Pending";
            }
            else
            {
                tempString = "";
            }
            return tempString;
        }

        public String getBuildDate(String asset)
        {
            String tempResult = "";
            PowerShell ps = PowerShell.Create();
            ps.AddScript("Invoke-Command -ComputerName " + asset + " {Get-ItemPropertyValue -Path HKLM:\\SOFTWARE\\FletcherBuilding -Name BuildDate}");
            Collection<PSObject> results = ps.Invoke();

            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject psObject in results)
            {
                stringBuilder.AppendLine(psObject.ToString());
            }
            tempResult = stringBuilder.ToString();

            return tempResult;
        }

        public String getBuildStatus(String asset)
        {
            String tempResult = "";
            PowerShell ps = PowerShell.Create();
            ps.AddScript("Invoke-Command -ComputerName " + asset + " {Get-ItemPropertyValue -Path HKLM:\\SOFTWARE\\FletcherBuilding -Name BuildStatus}");
            Collection<PSObject> results = ps.Invoke();

            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject psObject in results)
            {
                stringBuilder.AppendLine(psObject.ToString());
            }
            tempResult = stringBuilder.ToString();

            return tempResult;
        }
        public String getServiceRing(String asset)
        {
            String tempResult = "";
            PowerShell ps = PowerShell.Create();
            ps.AddScript("Invoke-Command -ComputerName " + asset + " {Get-ItemPropertyValue -Path HKLM:\\SOFTWARE\\FletcherBuilding -Name ServicingRing}");
            Collection<PSObject> results = ps.Invoke();

            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject psObject in results)
            {
                stringBuilder.AppendLine(psObject.ToString());
            }
            tempResult = stringBuilder.ToString();

            return tempResult;
        }

    }
}
