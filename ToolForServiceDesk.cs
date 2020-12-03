using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Management.Automation;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace TechTool
{
    class ToolForServiceDesk
    {

        public String getIPAddress(String asset, bool bool_UAT)
        {
            //To get Network Adapter names, use below command but IP address is in an array
            //Get-NetIPConfiguration -CimSession FBPF18RH65 | Where {$_.netadapter.status -ne 'Disconnected'} | select InterfaceDescription,IPv4Address,InterfaceAlias
            String tempString = "";
            String tempResult = "";
            String[] tempArray;
            String scriptString;
            PowerShell ps = PowerShell.Create();
            if(bool_UAT)
            {
                scriptString = "(Get-NetIPConfiguration -CimSession " + asset + " | Where {$_.netadapter.status -ne \'Disconnected\'}).ipv4address | select InterfaceAlias,IPv4Address";
            }else
            {
                scriptString = "([System.Net.Dns]::GetHostByName(\"" + asset + "\").AddressList[0]).IpAddressToString";
            }

            ps.AddScript(scriptString);
            //System.Windows.Forms.MessageBox.Show(scriptString);
            Collection <PSObject> results = ps.Invoke();

            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject psObject in results)
            {
                stringBuilder.AppendLine(psObject.ToString());
            }
            tempString = stringBuilder.ToString();
            //System.Windows.Forms.MessageBox.Show("TempString : " + tempString);
            tempResult = tempString;
            if (bool_UAT)
            {
                tempResult = "";
                tempArray = tempString.Split(new Char[] { '=', ';', '}' });
                for (int i = 1; i < tempArray.Length; i = i + 4)
                {
                    tempResult = tempResult + tempArray[i] + " : " + tempArray[i + 2] + "\r\n";
                }
            }
                
            return tempResult;
        }

        public String getHostName(String asset)
        {
            String tempHostName = "Unable to retrieve Host Name";
            try
            {
                IPHostEntry entry = Dns.GetHostEntry(asset);
                if (entry != null)
                {
                    tempHostName =  entry.HostName;
                }
            }
            catch (SocketException ex)
            {
                //unknown host or
                //not every IP has a name
                //log exception (manage it)
                Console.WriteLine(ex);
            }
            return tempHostName;
        }

        public void remoteDesktopViewer(String asset)
        {
            try
            {
                ProcessStartInfo info = new ProcessStartInfo("cmd", @"/c CmRcViewer.exe " + asset);
                Process p1 = Process.Start(info);
            }
            catch (Exception objException)
            {
                Console.WriteLine(objException);
            }
        }

        public String getUserLoggedIn(String asset)
        {
            String tempUserLoggedInString = "wmic /node:" + asset + " computersystem get username";
            String result = "Cannot find logged in user";
            String[] tempArray;
            try
            {
                ProcessStartInfo p = new ProcessStartInfo("cmd", "/c " + tempUserLoggedInString);
                String tempResult = "";

                p.RedirectStandardOutput = true;
                p.UseShellExecute = false;

                p.CreateNoWindow = true;

                Process proc = new Process();
                proc.StartInfo = p;
                proc.Start();

                tempResult = proc.StandardOutput.ReadToEnd();
                tempArray = tempResult.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                result = tempArray[1];
                result = result.Replace(" ", "").Replace("\n", "").Replace("\r", "");
            }
            catch (Exception objException)
            {
                Console.WriteLine(objException);
            }
            return result;
        }
        public void openRemoteCMD(String input)
        {
            try
            {
                ProcessStartInfo info = new ProcessStartInfo("cmd", @"/k AdditionalFiles\psexec.exe \\" + input + " cmd");
                Process p1 = Process.Start(info);
            }
            catch (Exception objException)
            {
                Console.WriteLine(objException);
            }
        }

        public void getMACAddress(String input)
        {
            //"cmd /c start cmd.exe /K getmac /s " + input
            String tempUserLoggedInString = "getmac /s " + input;

            ProcessStartInfo pro = new ProcessStartInfo("cmd.exe");
            String cmdCommand = @"getmac /s " + input;
            pro.Arguments = "/k" + cmdCommand;
            Process.Start(pro);

        }
        public void getSoftwareInstalled(String input)
        {
            try
            {
                ProcessStartInfo info = new ProcessStartInfo("cmd", @"/k AdditionalFiles\psinfo.exe \\" + input + " -s Applications");
                Process p1 = Process.Start(info);
            }
            catch (Exception objException)
            {
                Console.WriteLine(objException);
            }
        }

        public void getPCInfo(String input)
        {
            try
            {
                ProcessStartInfo info = new ProcessStartInfo("cmd", @"/k AdditionalFiles\psinfo.exe \\" + input + " -d");
                Process p1 = Process.Start(info);
            }
            catch (Exception objException)
            {
                Console.WriteLine(objException);
            }
        }

        public void getNetworkDrives(String asset, String userLoggedIn)
        {
            //for /f "delims=\ tokens=2" %A in ('reg query hku ^| findstr /i "S-1-5-" ^| findstr /v /i "_Classes"') do reg query HKU\%A\Network /s /v RemotePath	
            //reg query \\FBPF18RA5C\HKU\S-1-5-21-1062350055-1564878711-1247820936-135640\Network /s /v RemotePath
            String tempOutput = "";
            String tempSID = "";
            String[] tempArray;
            try
            {
                //convert username to SID
                tempArray = userLoggedIn.Split('\\');
                //"wmic useraccount where (name='" + usernameArray[1] + "' and domain='" + usernameArray[0] + "') get sid"
                ProcessStartInfo p = new ProcessStartInfo("cmd", "/c " + "wmic useraccount where (name='" + tempArray[1] + "' and domain='" + tempArray[0] + "') get sid");

                p.RedirectStandardOutput = true;
                p.UseShellExecute = false;

                p.CreateNoWindow = true;

                Process proc = new Process();
                proc.StartInfo = p;
                proc.Start();

                tempOutput = proc.StandardOutput.ReadToEnd();

                tempArray = tempOutput.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                tempSID = tempArray[1].Replace(" ", "");
                tempSID = tempSID.Replace("\r", "");

                //example : cmd /k reg query \\FBPF18RA5C\HKU\S-1-5-21-1062350055-1564878711-1247820936-135640\Network /s /v RemotePath
                p = new ProcessStartInfo("cmd", "/k " + "reg query \\\\" + asset + "\\HKU\\" + tempSID + "\\Network /s /v RemotePath");
                Process p1 = Process.Start(p);
                //var result = p1.StandardOutput.ReadToEnd();
            }
            catch (Exception objException)
            {
                Console.WriteLine(objException);
            }
        }

        public void getMappedPrinters(String asset, String userLoggedIn)
        {
            //for /f "delims=\ tokens=2" %A in ('reg query hku ^| findstr /i "S-1-5-" ^| findstr /v /i "_Classes"') do reg query HKU\%A\Network /s /v RemotePath	
            //reg query \\FBPF18RA5C\HKU\S-1-5-21-1062350055-1564878711-1247820936-135640\Network /s /v RemotePath
            String tempOutput = "";
            String tempSID = "";
            String[] tempArray;
            try
            {
                //convert username to SID
                tempArray = userLoggedIn.Split('\\');
                //"wmic useraccount where (name='" + usernameArray[1] + "' and domain='" + usernameArray[0] + "') get sid"
                ProcessStartInfo p = new ProcessStartInfo("cmd", "/c " + "wmic useraccount where (name='" + tempArray[1] + "' and domain='" + tempArray[0] + "') get sid");


                p.RedirectStandardOutput = true;
                p.UseShellExecute = false;

                p.CreateNoWindow = true;

                Process proc = new Process();
                proc.StartInfo = p;
                proc.Start();
                tempOutput = proc.StandardOutput.ReadToEnd();
                tempArray = tempOutput.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                tempSID = tempArray[1].Replace(" ", "");
                tempSID = tempSID.Replace("\r", "");

                //example : cmd /k reg query \\FBPF18RA5C\HKU\S-1-5-21-1062350055-1564878711-1247820936-135640\Printers\Connections
                p = new ProcessStartInfo("cmd", "/k " + "reg query \\\\" + asset + "\\HKU\\" + tempSID + "\\Printers\\Connections");
                Process p1 = Process.Start(p);
            }
            catch (Exception objException)
            {
                Console.WriteLine(objException);
            }
        }
        public ListBox getInstalledApps(ListBox listBox1, string asset)
        {
            //string remotemachine = "FBPF18RA5C";
            string remotemachine = asset;
            RegistryKey hive = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, remotemachine, RegistryView.Registry64);
            using (var key = hive.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"))
            {
                foreach (string skName in key.GetSubKeyNames())
                {
                    using (RegistryKey sk = key.OpenSubKey(skName))
                    {
                        try
                        {
                            if(sk.GetValue("DisplayName") != null && sk.GetValue("SystemComponent") == null && sk.GetValue("ParentKeyName") == null && sk.GetValue("ParentDisplayName") == null && sk.GetValue("ReleaseType") == null)
                            {
                                listBox1.Items.Add(sk.GetValue("DisplayName") + " (Version: " + sk.GetValue("DisplayVersion") + ")");
                            }
                            
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Exception : " + ex);
                        }
                    }
                }
            }
            hive = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, remotemachine, RegistryView.Registry64);
            using (var key = hive.OpenSubKey(@"Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"))
            {
                foreach (string skName in key.GetSubKeyNames())
                {
                    using (RegistryKey sk = key.OpenSubKey(skName))
                    {
                        try
                        {
                            if (sk.GetValue("DisplayName") != null && sk.GetValue("SystemComponent") == null && sk.GetValue("ParentKeyName") == null && sk.GetValue("ParentDisplayName") == null && sk.GetValue("ReleaseType") == null)
                            {
                                listBox1.Items.Add(sk.GetValue("DisplayName") + " (Version : " + sk.GetValue("DisplayVersion") + ")");
                            }

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Exception : " + ex);
                        }
                    }
                }
            }

            Console.WriteLine("Getting installed apps");
            string uninstallKey = @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall";
            //string uninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            /*using (RegistryKey rk = Registry.LocalMachine.OpenSubKey(uninstallKey))
            {
                foreach (string skName in rk.GetSubKeyNames())
                {
                    using (RegistryKey sk = rk.OpenSubKey(skName))
                    {
                        try
                        {
                            //Console.WriteLine(sk.GetValue("DisplayName"));
                            listBox1.Items.Add(sk.GetValue("DisplayName"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Exception : " + ex);
                        }
                    }
                }
                //label1.Text = listBox1.Items.Count.ToString();
                return listBox1;
            }*/
            return listBox1;
        }
    }
}
