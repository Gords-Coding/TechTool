using Microsoft.VisualBasic.FileIO;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices;
using System.Drawing;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TechTool
{
    public partial class Form1 : Form
    {

        ToolForServiceDesk tfsd = new ToolForServiceDesk();
        AccountManager am = new AccountManager();
        FBITHelpTool ht = new FBITHelpTool();
        RemoteRegistry rr = new RemoteRegistry();

        public Boolean steel_domain_bug = true; 
        public Boolean boolean_UAT = true;

        public Form1()
        {
            InitializeComponent();
            comboBox1.SelectedItem = "CompID/IP Add";
            checkBox1.Visible = boolean_UAT; 
        }


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.GetItemText(this.comboBox1.SelectedItem) == "Account Info")
            {
                panel2.Visible = false;
                panel3.Visible = true;
                panel4.Visible = false;
                panel5.Visible = false;
            }
            else if (comboBox1.GetItemText(this.comboBox1.SelectedItem) == "FB IT Help Tool Info")
            {
                panel2.Visible = false;
                panel3.Visible = false;
                panel4.Visible = true;
                panel5.Visible = false;
            }
            else if(comboBox1.GetItemText(this.comboBox1.SelectedItem) == "Remote Registry")
            {
                panel2.Visible = false;
                panel3.Visible = false;
                panel4.Visible = false;
                panel5.Visible = true;
            }
            else
            {
                panel2.Visible = true;
                panel3.Visible = false;
                panel4.Visible = false;
                panel5.Visible = false;
            }
        }


        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                pictureBox1_Click(sender, e);
            }
        }

        private void label5_Click(object sender, EventArgs e)
        {
            ProcessStartInfo pro = new ProcessStartInfo("cmd.exe");
            String openCmd = @"ping " + textBox1.Text;
            pro.Arguments = "/k" + openCmd;
            Process.Start(pro);
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if(textBox4.Text == "Successful")
            {
                button2.Enabled = true;
                button3.Enabled = true;
                button4.Enabled = true;
                button5.Enabled = true;
                button6.Enabled = true;
                button7.Enabled = true;
                button8.Enabled = true;
            }
            else
            {
                button2.Enabled = false;
                button3.Enabled = false;
                button4.Enabled = false;
                button5.Enabled = false;
                button6.Enabled = false;
                button7.Enabled = false;
                button8.Enabled = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            tfsd.openRemoteCMD(textBox1.Text);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            tfsd.getMACAddress(textBox1.Text);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            tfsd.getPCInfo(textBox1.Text);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //Need to add arraylist to list in form2
            //ArrayList softwareAL = new ArrayList();
            //Console.WriteLine("Hello World");

            Cursor.Current = Cursors.WaitCursor;
            ListBox listBox1 = new ListBox();
            Console.WriteLine("1");
            tfsd.getInstalledApps(listBox1, textBox1.Text);
            Console.WriteLine("2");
            var form2 = new Form2(listBox1);
            Console.WriteLine("3");
            form2.Show();
            Cursor.Current = Cursors.Default;
            Console.WriteLine("4");

            //this.Controls.Add(listBox1);

            /*ManagementObjectSearcher mos = new ManagementObjectSearcher("SELECT * FROM Win32_Product");
            foreach (ManagementObject mo in mos.Get())
            {
                //Console.WriteLine(mo["Name"]);
                softwareAL.Add(mo["Name"] + "\t" + mo["Caption"] + "\t" + mo["Vendor"] + "\t" + mo["Version"]);
            }
            softwareAL.Sort();*/
            /*Form2 f2 = new Form2();
            f2.ShowDialog();*/
            Console.WriteLine("End");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            tfsd.getNetworkDrives(textBox1.Text, tfsd.getUserLoggedIn(textBox1.Text));
            Cursor.Current = Cursors.Default;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            //int a = 25;
            //System.Windows.Forms.MessageBox.Show("$" + a);
            String input = textBox1.Text;
            Ping p1 = new Ping();
            if (comboBox1.SelectedItem.ToString() == "CompID/IP Add") {
                try
                {
                    //IPHostEntry hostInfo = Dns.Resolve(input);
                    //IPAddress[] address = hostInfo.AddressList;
                    //System.Windows.Forms.MessageBox.Show("Host name : " + hostInfo.HostName);
                    PingReply reply = p1.Send(input);
                    if (reply.Status == IPStatus.Success)
                    {
                        Console.WriteLine("Input Valid");
                        textBox4.ForeColor = Color.Green;
                        textBox4.Text = "Successful";
                    }
                    else
                    {
                        textBox4.ForeColor = Color.Red;
                        textBox4.Text = "Unsuccessful";
                    }

                    //if IP address, textbox3 get asset no
                    //and textBox4 set IP address given
                    if (textBox4.Text == "Successful")
                    {
                        //textBox2.Text = hostInfo.HostName;
                        textBox2.Text = tfsd.getHostName(input);
                        textBox5.Text = tfsd.getUserLoggedIn(input); 
                        textBox3.Text = tfsd.getIPAddress(input, checkBox1.Checked);
                    }
                    else
                    {
                        textBox2.Text = "";
                        textBox3.Text = "";
                        textBox4.ForeColor = Color.Red;
                        textBox4.Text = "Unsuccessful";
                        textBox5.Text = "";
                    }
                }
                catch (ArgumentNullException ex)
                {
                    Console.WriteLine("ArgumentNullException caught!!!");
                    Console.WriteLine("Source : " + ex.Source);
                    Console.WriteLine("Message : " + ex.Message);

                    textBox2.Text = "ArgumentNullException Caught!!!";
                    //textBox3.Text = ex.Message;
                    textBox4.ForeColor = Color.Red;
                    textBox4.Text = "Unsuccessful";
                    //textBox5.Text = "hostName is null.";

                }
                catch (ArgumentOutOfRangeException ex)
                {
                    Console.WriteLine("ArgumentOutOfRangeException Caught!!!");
                    Console.WriteLine("Source : " + ex.Source);
                    Console.WriteLine("Message : " + ex.Message);

                    textBox2.Text = "ArgumentOutOfRangeException Caught!!!";
                    //textBox3.Text = ex.Message;
                    textBox4.ForeColor = Color.Red;
                    textBox4.Text = "Unsuccessful";
                    //textBox5.Text = "The length of hostName is greater than 255 characters.";
                }
                catch (SocketException ex)
                {
                    Console.WriteLine("SocketException caught!!!");
                    Console.WriteLine("Source : " + ex.Source);
                    Console.WriteLine("Message : " + ex.Message);

                    textBox2.Text = "SocketException Caught!!!";
                    //textBox3.Text = ex.Message;
                    textBox4.ForeColor = Color.Red;
                    textBox4.Text = "Unsuccessful"; 
                    //textBox5.Text = "An error is encountered when resolving hostName";
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    textBox2.Text = "Caught!!!";
                    textBox4.ForeColor = Color.Red;
                    textBox4.Text = "Unsuccessful";
                }

            }
            else if (comboBox1.SelectedItem.ToString() == "FB IT Help Tool Info")
            {
                PingReply reply = p1.Send(input);
                if (reply.Status == IPStatus.Success)
                {
                    textBox11.Text = "Processing...";
                    textBox11.Text = ht.getComputerModel(input);
                    textBox13.Text = "Processing...";
                    textBox13.Text = ht.getComputerMake(input);
                    textBox15.Text = "Processing...";
                    textBox15.Text = ht.getComputerModelType(input);
                    textBox17.Text = "Processing...";
                    textBox17.Text = ht.getComputerOS(input);
                    textBox19.Text = "Processing...";
                    textBox19.Text = ht.getReleaseName(input);
                    textBox21.Text = "Processing...";
                    textBox21.Text = ht.getBuildDate(input);
                    textBox23.Text = "Processing...";
                    textBox23.Text = ht.getBuildStatus(input);
                    textBox25.Text = "Processing...";
                    textBox25.Text = ht.getServiceRing(input);
                    textBox27.Text = "Processing...";
                    textBox27.Text = ht.GetRestartPending(input);
                }
            }
            else if (comboBox1.SelectedItem.ToString() == "Remote Registry")
            {
                PingReply reply = p1.Send(input);
                if (reply.Status == IPStatus.Success)
                {
                    textBox28.Text = "\\\\" + ReplaceNonPrintableCharacters(textBox1.Text, "");
                    //Once connected to remote registry, give option to go in HKLM, HKU
                    button9.Enabled = true;
                    button10.Enabled = true;
                }
            }
            else
            {
                comboBox1.SelectedItem = "CompID/IP Add";
                pictureBox1_Click(sender, e);
            }
            Cursor.Current = Cursors.Default;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            tfsd.getMappedPrinters(textBox1.Text, tfsd.getUserLoggedIn(textBox1.Text));
            Cursor.Current = Cursors.Default;
        }

        private void label6_Click(object sender, EventArgs e)
        {
            String[] tempArray;
            String tempUsername = textBox5.Text;
            comboBox1.SelectedItem = "Account Info";
            tempArray = tempUsername.Split('\\');
            comboBox2.SelectedItem = tempArray[0];
            textBox6.Text = tempArray[1];
        }

        private void textBox6_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                pictureBox2_Click(sender, e);
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            label9.Text = "Processing...";
            textBox7.Text = "Processing...";
            textBox8.Text = "Processing...";
            textBox9.Text = "Processing...";
            Cursor.Current = Cursors.WaitCursor;
            string domainString = comboBox2.SelectedItem + "";

            //Reads DomainList.csv file to get DC name
            using (TextFieldParser parser = new TextFieldParser(@".\AdditionalFiles\DomainList.csv"))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                while (!parser.EndOfData)
                {
                    //Processing row until find match
                    string[] fields = parser.ReadFields();
                    if (fields[0].Equals(domainString))
                    {
                        domainString = fields[1];
                        break;
                    }
                }
            }

            var domainContext = new PrincipalContext(ContextType.Domain, domainString);
            try
            {
                //Declare user principal, if it goes through, AD account exist. If fails, AD account doesn't exist and will be caught
                var user = UserPrincipal.FindByIdentity(domainContext, textBox6.Text);
                DirectoryEntry entry = new DirectoryEntry("LDAP://" + user.DistinguishedName);
                
                label9.Text = "User found in local AD";
                label9.ForeColor = Color.Green;

                textBox7.Text = "Display Name : " + user.DisplayName;

                //Gets password last set and Expiring days
                if ((user.Enabled == true) && !(user.PasswordNeverExpires) && (user.LastPasswordSet != null))
                {
                    textBox8.Text = comboBox2.SelectedItem + " password last set: " + user.LastPasswordSet.Value.ToLocalTime() + ", " + am.getPasswordExpiringDays(domainString, textBox6.Text) + " days until expire";
                    //System.Windows.Forms.MessageBox.Show("Displayed textbox8 text");
                    if (entry.Properties["EmployeeID"].Value != null && !(comboBox2.SelectedItem.Equals("FBU")))
                    {
                        //If employeeID is not empty and combobox is not FBU
                        //Check password expiry date for FBU/mailbox account
                        try
                        {
                            var domainContext_2 = new PrincipalContext(ContextType.Domain, "UNWPDC2.fbu.com");
                            var fbuuser = UserPrincipal.FindByIdentity(domainContext_2, entry.Properties["EmployeeID"].Value.ToString());

                            textBox9.Text = "FBU/Mailbox password last set: " + user.LastPasswordSet.Value.ToLocalTime() + ", " + am.getPasswordExpiringDays("FBU", entry.Properties["EmployeeID"].Value.ToString()) + " days until expire";
                        }
                        catch
                        {
                            textBox9.Text = "Unable to find FBU account";
                        }
                        /*if (am.accountExist("FBU", am.getEmployeeID(domainString, textBox6.Text)))
                        {
                            //System.Windows.Forms.MessageBox.Show("FBU account Exists");
                            textBox9.Text = "FBU/Mailbox password last set: " + am.getPasswordLastSet("FBU", am.getEmployeeID(domainString, textBox6.Text)) + ", " + am.getPasswordExpiringDays("FBU", am.getEmployeeID(domainString, textBox6.Text)) + " days until expire";
                        }
                        else if (am.getEmployeeID(domainString, textBox6.Text) == "")
                        {
                            //System.Windows.Forms.MessageBox.Show("EmployeeID empty or combobox2 displays FBU");
                            textBox9.Text = "";
                        }
                        else
                        {
                            //System.Windows.Forms.MessageBox.Show("Else");
                            textBox9.Text = "Unable to find FBU account";
                        }*/
                    }
                    else
                    {
                        textBox9.Text = "";
                    }
                }
                else
                {
                    if (user.Enabled == false)
                    {
                        textBox8.Text = "Account Disabled";
                    }
                    else if (user.PasswordNeverExpires)
                    {
                        textBox8.Text = "Password Set To Never Expire";
                    } else if (user.LastPasswordSet == null)
                    {
                        textBox8.Text = "User must change password at next logon";
                    }
                    //System.Windows.Forms.MessageBox.Show("Account Enabled Status : " + am.getAccountEnabledStatus(comboBox2.SelectedItem + "", textBox6.Text));
                    //System.Windows.Forms.MessageBox.Show("Account Never Expires Status : " + am.getPasswordNeverExpiresStatus(comboBox2.SelectedItem + "", textBox6.Text));
                    textBox9.Text = "";
                }
            }
            catch
            {
                //System.Windows.Forms.MessageBox.Show("Unable to process user");
                label9.Text = "Account does not Exist/Unable to process user";
                label9.ForeColor = Color.Red;

                textBox7.Text = "N/A";
                textBox8.Text = "N/A";
                textBox9.Text = "N/A";
            }
            
            //System.Windows.Forms.MessageBox.Show(user.LastPasswordSet + "");
            

            /*if (steel_domain_bug && comboBox2.GetItemText(this.comboBox2.SelectedItem) == "FCL-FCSS") {
                domainString = "fsl-adc01.fsl.fb.net.nz";
            } */
            

            /*if (label9.Text == "User found in local AD")
            {
                if (am.getAccountEnabledStatus(domainString, textBox6.Text) && !(am.getPasswordNeverExpiresStatus(domainString, textBox6.Text)))
                {
                    //System.Windows.Forms.MessageBox.Show(textBox6.Text + " Account Enabled and Password Will Expire");
                    
                    textBox8.Text = comboBox2.SelectedItem + " password last set: " + am.getPasswordLastSet(domainString, textBox6.Text) + ", " + am.getPasswordExpiringDays(domainString, textBox6.Text) + " days until expire";
                    //System.Windows.Forms.MessageBox.Show("Displayed textbox8 text");
                    if (am.getEmployeeID(domainString, textBox6.Text) != "") 
                    {
                        //System.Windows.Forms.MessageBox.Show("Combobox2 does not display FBU and EmployeeID is not empty");
                        if (am.accountExist("FBU", am.getEmployeeID(domainString, textBox6.Text)))
                        {
                            //System.Windows.Forms.MessageBox.Show("FBU account Exists");
                            textBox9.Text = "FBU/Mailbox password last set: " + am.getPasswordLastSet("FBU", am.getEmployeeID(domainString, textBox6.Text)) + ", " + am.getPasswordExpiringDays("FBU", am.getEmployeeID(domainString, textBox6.Text)) + " days until expire";
                        }
                        else if (am.getEmployeeID(domainString, textBox6.Text) == "")
                        {
                            //System.Windows.Forms.MessageBox.Show("EmployeeID empty or combobox2 displays FBU");
                            textBox9.Text = "";
                        }
                        else
                        {
                            //System.Windows.Forms.MessageBox.Show("Else");
                            textBox9.Text = "Unable to find FBU account";
                        }
                    }else
                    {
                        textBox9.Text = "";
                    }
                }
                else
                {
                    if (!am.getAccountEnabledStatus(domainString, textBox6.Text))
                    {
                        textBox8.Text = "Account Disabled";
                    }
                    else if (am.getPasswordNeverExpiresStatus(domainString, textBox6.Text))
                    {
                        textBox8.Text = "Password Set To Never Expire";
                    }
                    //System.Windows.Forms.MessageBox.Show("Account Enabled Status : " + am.getAccountEnabledStatus(comboBox2.SelectedItem + "", textBox6.Text));
                    //System.Windows.Forms.MessageBox.Show("Account Never Expires Status : " + am.getPasswordNeverExpiresStatus(comboBox2.SelectedItem + "", textBox6.Text));
                    textBox9.Text = "";
                }
            }
            else
            {
                textBox7.Text = "N/A";
                textBox8.Text = "N/A";
                textBox9.Text = "N/A";
            }*/

            Cursor.Current = Cursors.Default;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox4.Text == "Successful")
            {
                tfsd.remoteDesktopViewer(textBox1.Text);
            }
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

        private void button9_Click(object sender, EventArgs e)
        {
            textBox28.Text = "\\\\" + textBox1.Text + "\\HKLM";
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            String[] tempArray = rr.regQuery(textBox28.Text);
            for(int i = 1; i < tempArray.Length; i++)
            {
                if (!tempArray[i].ToString().Contains("    "))
                {
                    listBox1.Items.Add(tempArray[i].ToString());
                }
            }
            
            tempArray = rr.regQueryValue(textBox28.Text);
            for (int i = 1; i < tempArray.Length; i++)
            {
                listBox2.Items.Add(tempArray[i].ToString());
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            textBox28.Text = "\\\\" + textBox1.Text + "\\HKU";
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            String[] tempArray = rr.regQuery(textBox28.Text);
            for (int i = 1; i < tempArray.Length; i++)
            {
                if (!tempArray[i].ToString().Contains("    "))
                {
                    listBox1.Items.Add(tempArray[i].ToString());
                }
            }

            tempArray = rr.regQueryValue(textBox28.Text);
            for (int i = 1; i < tempArray.Length; i++)
            {
                listBox2.Items.Add(tempArray[i].ToString());
            }
        }

        private void textBox28_TextChanged(object sender, EventArgs e)
        {
            textBox28.Text.Replace("HKEY_USERS", "HKU");
            textBox28.Text.Replace("HKEY_LOCAL_MACHINE", "HKLM");
        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            textBox28.Text = "\\\\" + textBox1.Text + "\\" + listBox1.SelectedItem.ToString();
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            String[] tempArray = rr.regQuery(textBox28.Text);
            for (int i = 1; i < tempArray.Length; i++)
            {
                if (!tempArray[i].ToString().Contains("    "))
                {
                    listBox1.Items.Add(tempArray[i].ToString());
                }
            }

            tempArray = rr.regQueryValue(textBox28.Text);
            for (int i = 1; i < tempArray.Length; i++)
            {
                listBox2.Items.Add(tempArray[i].ToString());
            }
        }


        private void checkBox1_CheckedStateChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                //System.Windows.Forms.MessageBox.Show("Tickbox ticked");
            }
            else
            {
                //System.Windows.Forms.MessageBox.Show("Tickbox unticked");
            }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            Clipboard.SetText("Asset no : " + textBox2.Text + "\nUsername : " + textBox5.Text + "\nIP : " + textBox3.Text);
            label10.Text = "Copied!";
            delay(2000, () => label10.Text = "");
        }
        public void delay(int delay, Action action)
        {
            Timer timer = new Timer();
            timer.Interval = delay;
            timer.Tick += (s, e) => {
                action();
                timer.Stop();
            };
            timer.Start();
        }
    }
}
