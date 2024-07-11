using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
//SnmpSharpNet is a 3rd party library that allows us to connect to the racks. 
using SnmpSharpNet;
using System.Net;
using System.Net.NetworkInformation;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Diagnostics;
namespace TrippLite_GUI
{//need to fix it so that multiple deletes occur only can do one at a time right now
    public partial class Main : Form
    {
        //private Excel.Workbook theWorkbook;
        //private Excel.Sheets sheets;
        //private Excel.Worksheet worksheet;
        //private Excel.Range range;
        private string[][] strArray;
        private Boolean racTrue;
        private Excel.Application ExcelObj = new Excel.Application();
        private bool search1;
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        /// <Summary>
        /// Begins the program. It reads in the spreadsheet values that have both names and rack numbers. Then it proceeds
        /// to fill in the treeview1 node names to match what device is in that area. Finally when it completes filling in the 
        /// spreadsheet it opens up the GUI window. 
        /// 
        /// If anyone ever needs to change the gui because the rack's ip changed then you need to change all the places in the gui
        /// that have the integer 89, 90 or any where in the hundreds locations. Since those statements are dictating which rack has 
        /// what ip. 
        /// @author Fast Interop
        /// 7/20/2016
        /// </summary>
//----------------------------------------Constructor----------------------------------------------------------------------------------------------
        public Main()
        {
            // INSTANTIATE strArray
            strArray = new string[408][];
            InitializeComponent();
            //make splash screen appear
            SplashScreen splash = new SplashScreen();
            splash.Show();
            uint excelId = 0;
            // CHECK IF EXCEL IS AVAIABLE,
            // if not, close the program as it needs excel.
            if (ExcelObj == null)
            {
                MessageBox.Show("Unable to connect to Microsoft Excel! Terminating program.", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Windows.Forms.Application.Exit();
            }
            search1 = false;
            //opens the excel book and selects the first sheet
            var book = ExcelObj.Workbooks;
            var theWorkbook = book.Open(
                @"\\fox\raid\ethernets\Interop\Interop Randomizer\List of Link Partners.xls", 0, true, 5,
                    "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false,
                    0, true);
            //gets all the sheets in the workbook
            var sheets = theWorkbook.Worksheets;
            int l = 1;
            int loc = 0;
            GetWindowThreadProcessId(new IntPtr(ExcelObj.Hwnd), out excelId);
            //traverses through the excel worksheets until there are no more pages
            while (l <= sheets.Count)
            {
                var worksheet = (Excel.Worksheet)sheets.get_Item(l);
                int workSheetCount = sheets.Count;
                int i = 2;
                //creates the range of cells to look at
                var myRange = worksheet.get_Range("A" + i);
                int count = 0;
                //uses the myRange value to know when it is at the end of the spreadsheet page. 
                while (myRange.Value2 != null)
                {
                    Excel.Range iolRange = worksheet.get_Range("E" + i);
                    //if-else statement that figures out if the device is part of the racks and calculates the last 
                    // location to be filled. Makes sure there is no empty space in the array list.
                    if (iolRange.Value2 == null)
                    {
                        //if the cell is blank it does this
                        count += 1;
                        i += 1;
                        myRange = worksheet.get_Range("A" + i);
                    }
                    else
                    {
                        //if the cell has a value it performs these functions to insert the values into the correct location
                        //in the rack arrays/treeview object
                        myRange = worksheet.get_Range("A" + i);
                        var range = worksheet.get_Range("A" + i, "F" + i);
                        System.Array myvalues = (System.Array)range.Cells.Value;
                        strArray[i - (2 + count) + loc] = ConvertToStringArray(myvalues);
                        if (strArray[i - (2 + count) + loc][0] != "")
                        {
                            //stores the name of the worksheet
                            strArray[i - (2 + count) + loc][strArray[0].Length - 3] = worksheet.Name;
                            //stores the value of the worksheet the item was found on so that we can make edits later
                            strArray[i - (2 + count) + loc][strArray[0].Length - 2] = (string)l.ToString();
                            //stores the value of the original row of the spreadsheet the item was found
                            strArray[i - (2 + count) + loc][strArray[0].Length - 1] = (string)i.ToString();
                        }
                        string rac = strArray[i - (2 + count) + loc][3];
                        string ol = strArray[i - (2 + count) + loc][4];
                        //fills in the names on the treeview1 object.
                        //splits the values of the outlets at the & symbol
                        string[] ands = ol.Split('&');
                        if (ands.Length > 1)
                        {
                            //meant for things that hvae more than 2 outlets
                            for (int h = 0; h < ands.Length; h++)
                            {
                                //convert each of the string in the array from string to int to fill in the treeview object correctly
                                int r;
                                int o;
                                bool isOutlet = int.TryParse(ands[h], out o);
                                bool isRack = int.TryParse(rac, out r);
                                //decides the location of the treeview node to fill in and what to fill in there.
                                if (isRack && isOutlet)
                                    treeView1.Nodes[r - 1].Nodes[o - 1].Text = o + ". " + strArray[i - (2 + count) + loc][0];
                            }
                        }
                        //meant for things with only one outlet
                        else
                        {
                            //r = rack number
                            int r;
                            //o = outlet number
                            int o;
                            bool isOutlet = int.TryParse(ol, out o);
                            bool isRack = int.TryParse(rac, out r);
                            //decides the location of the treeview node to fill in and what to fill in there.
                            if (isRack && isOutlet)
                                treeView1.Nodes[r - 1].Nodes[o - 1].Text = o + ". " + strArray[i - (2 + count) + loc][0];
                        }
                        i += 1;
                        myRange = worksheet.get_Range("A" + i);
                        CloseExcel(range);
                    }
                    CloseExcel(iolRange);
                }
                loc = loc + i - (2 + count);
                l += 1;
                CloseExcel(myRange);
            }
            if (loc < 408)
            {
                for (int i = loc; i < 408; i++)
                {
                    string[] s = new string[1];
                    s[0] = "";
                    strArray[i] = s;
                }
            }
            //close all excel objects that are currently open
            theWorkbook.Close(false);
            ExcelObj.Quit();
            CloseExcel(sheets);
            CloseExcel(book);
            CloseExcel(theWorkbook);
            CloseExcel(ExcelObj);
            //the following try catch and process tracking was a solution from Jordy "Kaiwa" Ruiter found on
            //www.codeproject.com/Questions/74980/Close-Excel-Process-with-Interop
            try
            {
                if (excelId != 0)
                {
                    Process excel = Process.GetProcessById((int)excelId);
                    excel.CloseMainWindow();
                    excel.Refresh();
                    excel.Kill();
                }
            }
            catch
            {
                //process was already killed
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            //create a new thread list to prepare for multithreading
            List<Thread> tlist = new List<Thread>();
            for (int k = 90; k < 107; k++)
            {
                racTrue = false;
                string ips = "132.177.122." + k.ToString();
                bool isConnected = PingHost(ips);
                if (isConnected == true)
                {

                    Dictionary<Tuple<int, int>, string> map = new Dictionary<Tuple<int, int>, string>();
                    //for loop for each of the outlets in the rack 
                    for (int p = 1; p < 25; p++)
                    {
                        Tuple<int, int> tup = new Tuple<int, int>(k, p);
                        map.Add(tup, ips);
                        if (treeView1.Nodes[k - 90].Nodes[p - 1].ForeColor == Color.LimeGreen)
                        {
                            racTrue = true;
                        }
                    }

                    //makes a new thread for each of the racks
                    Thread t = new Thread(() => doWork(map));
                    //adds the thread to the list
                    tlist.Add(t);
                    t.Start();



                }
            }
            for (int i = 0; i < tlist.Count; i++)
            {
                tlist[i].Join();
            }
            //get rid of the splash screen
            splash.Hide();
        }
        /// <summary>
        /// Does the work of populating racks in background. Allows to run multiple strings at once by using multithreading 
        /// </summary>
        /// <param name="map"></param>
//--------------------------------------------------------doWork-------------------------------------------------------------------
        private void doWork(Dictionary<Tuple<int, int>, string> map)
        {
            try
            {


                int result = 0;
                int k = 0;
                List<int> list = new List<int>();
                //for each of the racks that populate the multithread list
                foreach (KeyValuePair<Tuple<int, int>, string> entry in map)
                {
                    //check the status of a specific rack/outlet in that rack
                    result = snmpGet(entry.Key.Item2, entry.Value);
                    if (result == -1)
                    {
                        break;
                    }
                    //set specific outlet to k
                    k = entry.Key.Item1;
                    //set tree node for rack green if true. 
                    if (racTrue == true)
                        treeView1.Nodes[k - 90].ForeColor = Color.LimeGreen;
                    else
                        treeView1.Nodes[k - 90].ForeColor = Color.LimeGreen;
                }
            }
            catch (SnmpException e)
            {
                Console.WriteLine("There was and error" + e);
            }

        }
        /// <summary>
        /// Converts "values" to a string array
        /// </summary>
        /// <param name="values">Input Array</param>
        /// <returns>Output Array</returns>
//--------------------------Convert to String array-------------------------------------------------------------------------------
        string[] ConvertToStringArray(System.Array values)
        {

            // create a new string array
            string[] theArray = new string[values.Length + 3];


            // loop through the 2-D System.Array and populate the 1-D String Array
            for (int i = 1; i <= values.Length; i++)
            {
                if (i == 3)
                {
                    //do nothing
                }
                else
                {
                    if (i < 3)
                    {
                        if (values.GetValue(1, i) == null)
                            theArray[i - 1] = "";
                        else
                            theArray[i - 1] = (string)values.GetValue(1, i).ToString();
                    }
                    else
                    {
                        if (values.GetValue(1, i) == null)
                            theArray[i - 2] = "";
                        else
                            theArray[i - 2] = (string)values.GetValue(1, i).ToString();
                    }
                }
            }

            return theArray;
        }
        /// <summary>
        /// controls the racks power. Double click to turn on.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        /*
//------------------------------------------------Node mouse double click----------------------------------------------------------
        private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            //controls the toggle to turn on/off the outlets in the rack. 
            if (treeView1.SelectedNode != null & treeView1.SelectedNode.Text.Contains("Rack"))
            {

            }
            else
            {
                int ips = treeView1.SelectedNode.Parent.Index;
                int ol = treeView1.SelectedNode.Index;
                string ip = "132.177.122.";
                ips = ips + 90;
                ol = ol + 1;
                ip = ip + ips;
                //THIS PROTECTS THE INTERNET PATCH IN RACK1 FROM BEING TURNED OFF DO NOT CHANGE UNLESS YOU NEED TO CHANGE THE IP OF ALL THE RACKS
                if (ip == "132.177.122.90" && (ol == 18 || ol == 19 || ol == 20 || ol == 21 || ol == 22 || ol == 23 || ol == 24))
                {
                    MessageBox.Show("Power control for internet switches not allowed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    //calls the function that actually toggles the power in the rack. 
                    snmpSet(ip, ol);
                }
            }
        }
        */
        /// <summary>
        /// returns a boolean value that informs the program if the outlet is on or off.
        /// Talks directly to the racks using snmp.
        /// </summary>
        /// <param name="outlet"></param>
        /// <param name="ipValue"></param>
        /// <returns></returns>
//---------------------------------------------snmpGet-----------------------------------------------------------------------------
        private int snmpGet(int outlet, string ipValue)
        {

            string[] rackIP = ipValue.Split('.');
            int racknum;
            int.TryParse(rackIP[3], out racknum);
            //If rack numbers change you must change the racknum equation below since it is meant for the rack numbers from 90 - 105. 
            //This calculates which rack we are working on, ex. if rack 90 it correlates with rack 1
            racknum = racknum - 89;
            // Prepare target ( Rack )
            UdpTarget target = new UdpTarget((IPAddress)new IpAddress(ipValue));
            // Create a SET PDU
            Pdu pdu = new Pdu(PduType.Get);
            // Set sysLocation.0 to a new string
            pdu.VbList.Add(new Oid("1.3.6.1.4.1.850.100.1.10.2.1.2." + outlet.ToString()));
            // Set Agent security parameters
            AgentParameters aparam = new AgentParameters(SnmpVersion.Ver2, new OctetString("tripplite"));
            // Response packet
            SnmpV2Packet response;
            try
            {
                // Send request and wait for response
                response = target.Request(pdu, aparam) as SnmpV2Packet;
            }
            catch (Exception ex)
            {
                // If exception happens, it will be returned here
                Console.WriteLine("Error");
                target.Close();
                return -1;
            }

            //if it returns 2 then the outlet is considered on and therefore should be colored green to represent on
            if (response.Pdu[0].Value.ToString() == "2")
            {
                //set the specific outlet name to green in the tree view
                treeView1.Nodes[racknum - 1].Nodes[outlet - 1].ForeColor = Color.LimeGreen;
                return 1;
            }
            //else it is off and the name of the outlet should be red to signify that it is off
            else
            {
                //set specific outlet name to red in the treeview 
                treeView1.Nodes[racknum - 1].Nodes[outlet - 1].ForeColor = Color.Red;
                return 0;
            }



        }
        /// <summary>
        /// changes the value of the snmp to either turn on or turn off the rack depending upon the previous status. 
        /// </summary>
        /// <param name="ipValue"></param>
        /// <param name="outlet"></param>
//--------------------------------------------snmpSet()---------------------------------------------------------------------------
        private void snmpSet(string ipValue, int outlet, int on_off)
        {
            int count = 0;
            //splits up the ip value to get the last 3 digits
            string[] ip = ipValue.Split('.');
            int s;
            //turns the last 3 digits of the last 3 digits of the ip value from string to int
            int.TryParse(ip[3], out s);
            //subtract 90 to determine what node to deal with
            s = s - 90;

            int firstValue = 0;
            //get the outlet number from the treeview node name
            string place = treeView1.Nodes[s].Nodes[outlet - 1].Text.Substring(0, 2);
            int fPlace;
            //turn the outlet number into an int
            int.TryParse(place, out fPlace);
            //used to check if an object occupys more than one outlet in a rack. 
            for (int a = 0; a < 24; a++)
            {
                //get the outlet number from the front of the name
                string listPlace = treeView1.Nodes[s].Nodes[outlet - 1].Text.Substring(0, 2);
                int fListPlace;
                //convert the outlet number from string to int
                int.TryParse(listPlace, out fListPlace);
                //check for things below 10 because it changes the location of substring that the name starts at
                if (fPlace < 10 && fListPlace < 10)
                {
                    if (treeView1.Nodes[s].Nodes[outlet - 1].Text.Substring(3, treeView1.Nodes[s].Nodes[outlet - 1].Text.Length - 4)
                        == treeView1.Nodes[s].Nodes[a].Text.Substring(3, treeView1.Nodes[s].Nodes[a].Text.Length - 4))
                    //Argument out of bounds
                    {
                        if (count == 0)
                        {
                            firstValue = a + 1;
                        }
                        count += 1;

                    }
                }
                //check for larger than 9 because it changes the location of where the substring for the name of the device starts
                else if (fPlace > 9 && fListPlace > 9)
                {
                    if (treeView1.Nodes[s].Nodes[outlet - 1].Text.Substring(4, treeView1.Nodes[s].Nodes[outlet - 1].Text.Length - 4)
                       == treeView1.Nodes[s].Nodes[a].Text.Substring(4, treeView1.Nodes[s].Nodes[a].Text.Length - 4))
                    //Argument out of bounds
                    {
                        if (count == 0)
                        {
                            firstValue = a + 1;
                        }
                        count += 1;

                    }
                }
                //to check in the area of overlap if it is both in low and high areas
                else if (fPlace > 9 && fListPlace < 10)
                {
                    if (treeView1.Nodes[s].Nodes[outlet - 1].Text.Substring(3, treeView1.Nodes[s].Nodes[outlet - 1].Text.Length - 4)
                       == treeView1.Nodes[s].Nodes[a].Text.Substring(4, treeView1.Nodes[s].Nodes[a].Text.Length - 4))
                    //Argument out of bounds
                    {
                        if (count == 0)
                        {
                            firstValue = a + 1;
                        }
                        count += 1;

                    }
                }
                //general check at the end
                else
                {
                    if (treeView1.Nodes[s].Nodes[outlet - 1].Text.Substring(4, treeView1.Nodes[s].Nodes[outlet - 1].Text.Length - 4)
                       == treeView1.Nodes[s].Nodes[a].Text.Substring(3, treeView1.Nodes[s].Nodes[a].Text.Length - 4))
                    //Argument out of bounds
                    {
                        if (count == 0)
                        {
                            firstValue = a + 1;
                        }
                        count += 1;

                    }
                }
            }
            //check status of the outlets that we found
            //on_off acting as a boolean
            //on_off = snmpGet(outlet, ipValue);
            Console.WriteLine("Int: " + on_off);
            // Prepare target
            UdpTarget target = new UdpTarget((IPAddress)new IpAddress(ipValue));
            // Create a SET PDU
            Pdu pdu = new Pdu(PduType.Set);
            //check if device has more than one outlet && check to make sure not "Empty" otherwise need to catch so every "Empty" outlet doesnt turn on
            if (count > 1 && !treeView1.Nodes[s].Nodes[outlet - 1].ToString().Contains("Empty"))
            {
                //used to help to catch empty values
                if (fPlace < 10)
                {
                    //used to catch empty values
                    if (treeView1.Nodes[s].Nodes[outlet - 1].Text.Substring(3, treeView1.Nodes[s].Nodes[outlet - 1].Text.Length - 4) != "Empty")
                    {
                        int tempFirst = firstValue;
                        //turn on the array of outlets since they will be right next to each other.
                        for (int p = 0; p < count; p++)
                        {
                            //check if the device is on or off to start... if on turn off
                            if (on_off == 1)
                            {
                                label3.Text = "Device Power Status: Powering Off...";
                                //tells the pdu the value one which tells it to shut off the specific outlet. 
                                pdu.VbList.Add(new Oid("1.3.6.1.4.1.850.100.1.10.2.1.4." + firstValue.ToString()), new Integer32(1));
                                System.Threading.Thread.Sleep(1000);
                                firstValue = firstValue + 1;
                            }
                            else if (on_off == 3)
                            {
                                Console.WriteLine("Old: " + firstValue);
                                Console.WriteLine("Off");
                                label3.Text = "Device Power Status: Power Cycling...";
                                //tells the pdu the value one which tells it to shut off the specific outlet. 
                                pdu.VbList.Add(new Oid("1.3.6.1.4.1.850.100.1.10.2.1.4." + firstValue.ToString()), new Integer32(1));
                                System.Threading.Thread.Sleep(5000);
                                firstValue = firstValue + 1;
                            }
                            //else turn the device from off to on
                            else
                            {
                                label3.Text = "Device Power Status: Powering On...";
                                //tells the pdu 2 which tells the pdu to power on the device
                                pdu.VbList.Add(new Oid("1.3.6.1.4.1.850.100.1.10.2.1.4." + firstValue.ToString()), new Integer32(2));
                                System.Threading.Thread.Sleep(1000);
                                firstValue = firstValue + 1;
                            }
                        }
                        if (on_off == 3)
                        {
                            Thread.Sleep(20000);
                            firstValue = tempFirst;
                            Console.WriteLine("new: " + firstValue);
                            for (int p = 0; p < count; p++)
                            {
                                Console.WriteLine("on");
                                label3.Text = "device power status: power cycling...";
                                pdu.VbList.Add(new Oid("1.3.6.1.4.1.850.100.1.10.2.1.4." + firstValue.ToString()), new Integer32(2));
                                System.Threading.Thread.Sleep(5000);
                                firstValue = firstValue + 1;
                            }
                        }
                    }
                }
                else
                {
                    //check for "Empty" value
                    if (treeView1.Nodes[s].Nodes[outlet - 1].Text.Substring(4, treeView1.Nodes[s].Nodes[outlet - 1].Text.Length - 4) != "Empty")
                    {
                        int tempFirst = firstValue;
                        //loop through all of the outlets with the same name
                        for (int p = 0; p < count; p++)
                        {
                            //if they are on turn them off.
                            if (on_off == 1)
                            {
                                label3.Text = "Device Power Status: Powering Off...";
                                pdu.VbList.Add(new Oid("1.3.6.1.4.1.850.100.1.10.2.1.4." + firstValue.ToString()), new Integer32(1));
                                System.Threading.Thread.Sleep(1000);
                                firstValue = firstValue + 1;
                            }
                            else if (on_off == 3)
                            {
                                label3.Text = "Device Power Status: Power Cycling...";
                                pdu.VbList.Add(new Oid("1.3.6.1.4.1.850.100.1.10.2.1.4." + firstValue.ToString()), new Integer32(1));
                                System.Threading.Thread.Sleep(5000);
                                firstValue = firstValue + 1;
                            }
                            //if they are off turn them on
                            else
                            {
                                label3.Text = "Device Power Status: Powering On...";
                                pdu.VbList.Add(new Oid("1.3.6.1.4.1.850.100.1.10.2.1.4." + firstValue.ToString()), new Integer32(2));
                                System.Threading.Thread.Sleep(1000);
                                firstValue = firstValue + 1;
                            }
                        }
                        if (on_off == 3)
                        {
                            firstValue = tempFirst;
                            for (int p = 0; p < count; p++)
                            {
                                label3.Text = "Device Power Status: Power Cycling...";
                                pdu.VbList.Add(new Oid("1.3.6.1.4.1.850.100.1.10.2.1.4." + firstValue.ToString()), new Integer32(2));
                                System.Threading.Thread.Sleep(5000);
                                firstValue = firstValue + 1;
                            }
                        }
                    }
                }
            }
            else
            {
                //if only one outlet or is an empty outlet this occurs so that it only changes the status of one outlet.
                if (on_off == 1)
                {
                    label3.Text = "Device Power Status: Powering Off...";
                    pdu.VbList.Add(new Oid("1.3.6.1.4.1.850.100.1.10.2.1.4." + outlet.ToString()), new Integer32(1));
                    System.Threading.Thread.Sleep(1000);
                }
                else if (on_off == 3)
                {
                    label3.Text = "Device Power Status: Power Cycling...";
                    pdu.VbList.Add(new Oid("1.3.6.1.4.1.850.100.1.10.2.1.4." + outlet.ToString()), new Integer32(1));
                    System.Threading.Thread.Sleep(1000);
                    pdu.VbList.Add(new Oid("1.3.6.1.4.1.850.100.1.10.2.1.4." + outlet.ToString()), new Integer32(2));
                    System.Threading.Thread.Sleep(1000);
                }
                else
                {
                    label3.Text = "Device Power Status: Powering On...";
                    pdu.VbList.Add(new Oid("1.3.6.1.4.1.850.100.1.10.2.1.4." + outlet.ToString()), new Integer32(2));
                    System.Threading.Thread.Sleep(1000);
                }
            }
            AgentParameters aparam = new AgentParameters(SnmpVersion.Ver2, new OctetString("tripplite"));
            // Response packet
            SnmpV2Packet response;
            try
            {
                // Send request and wait for response
                response = target.Request(pdu, aparam) as SnmpV2Packet;
            }
            catch (Exception ex)
            {
                // If exception happens, it will be returned here
                target.Close();
                return;
            }
            // Make sure we received a response
            if (response == null)
            {

            }
            else
            {
                // Check if we received an SNMP error from the agent
                if (response.Pdu.ErrorStatus != 0)
                {

                }
            }
        }
        /// <summary>
        /// Ping Host checks if the host is connected on the correct ip. 
        /// If it is not connected will skip and move to the next one. 
        /// Shows up green if connected/ black if not.
        /// </summary>
        /// <param name="nameOrAddress"></param>
        /// <returns></returns>
//---------------------------------------------Ping Host---------------------------------------------------------------------------
        private bool PingHost(string nameOrAddress)
        {
            bool pingable = false;
            Ping pinger = new Ping();
            try
            {
                //try to ping specifc IP address of the rack
                PingReply reply = pinger.Send(nameOrAddress);
                pingable = reply.Status == IPStatus.Success;
            }
            catch (PingException)
            {
                // Discard PingExceptions and return false;
            }
            return pingable;
        }


        /// <summary>
        /// closes the excel document that was used for read-in.
        /// </summary>
        /// <param name="excel"></param>
        private void CloseExcel(object excel)
        {
            try
            {
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(excel) > 0) ;
            }
            catch (Exception ex)
            {

            }
            finally
            {
                excel = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }


        /// <summary>
        /// Tells the program what to do when a new tree node is selected (Highlighted).
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

            if (!search1 && treeView1.SelectedNode != null)
            {
                int count = 0;
                //checks to see if the device is on or off
                if (treeView1.SelectedNode.ForeColor == Color.Red || treeView1.SelectedNode.ForeColor == Color.LimeGreen)
                {
                    if (!search1)
                        //updates information on the right hand side of the window
                        updateField();
                    for (int j = 0; j < 16; j++)
                    {
                        //check to see how many tree nodes are expanded
                        if (treeView1.Nodes[j].IsExpanded)
                            count += 1;
                    }
                }
                //if more that 2 tree nodes expanded close the ones that people try to open after that
                //otherwise the program will run super slow because it updates subnodes whenever the head nodes
                //are expanded. 
                if (count > 2)
                {
                    e.Node.Collapse();
                }
            }
        }
        /// <summary>
        /// Function for the search button. 
        /// Utilizes the IOL search method when pressed.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
//---------------------------------------------Button Click----------------------------------------------------------------------
        private void button1_Click(object sender, EventArgs e)
        {
            search1 = true;
            //search for the IOL number entered into the search bar. 
            iolSearch();
            search1 = false;
        }
        /// <summary>
        /// iolSearch, looks through the array list to search for the iol number that is listed in the search bar.
        /// 
        /// </summary>
//-------------------------------------------------IOL Search-----------------------------------------------------------------------
        private void iolSearch()
        {
            string search = textBox1.Text;
            int iol;
            int iolnum = 0;
            int rack;
            int outlet;
            //in the case that the person actually filled in an IOL number
            if (search != "" || search != null)
            {
                //turn the inputed number from a string to an int
                int.TryParse(search, out iol);
                //close all expanded tree nodes
                treeView1.CollapseAll();
                try
                {
                    //run through the list and try to find the specified IOL number
                    for (int i = 0; i < strArray.Length - 1; i++)
                    {
                        int.TryParse(strArray[i][1], out iolnum);
                        int.TryParse(strArray[i][3], out rack);
                        string[] ands = strArray[i][4].Split('&');
                        //As you traverse the list if multiple outlets appear seperate them and take the first input. 
                        if (ands.Length > 1)
                        {
                            //assign first outlet to outlet
                            int.TryParse(ands[0], out outlet);
                        }
                        else
                        {
                            //assign first outlet to outlet
                            int.TryParse(strArray[i][4], out outlet);
                        }
                        if (iol == iolnum)
                        {
                            //when the IOL number is found expand the rack that it is currently in and have the selector
                            //highlight the name of the device in treeview 
                            treeView1.Nodes[rack - 1].Expand();
                            treeView1.SelectedNode = treeView1.Nodes[rack - 1].Nodes[outlet - 1];
                            treeView1.Focus();
                            updateField();
                            search1 = false;
                            return;

                        }
                        
                    }

                }
                //in the case of the IOL number not existing in the list
                catch (Exception e)
                {
                    System.Windows.Forms.MessageBox.Show("IOL Number Not Found");
                }
            }
        }
        /// <summary>
        /// snmpHalt is the function that is called to stop devices. 
        /// It is very similar to snmpSet except it will not toggle devices on, only turn them off. 
        /// Needed to "Halt" the racks (turn off all of the devices inside of the racks. 
        /// </summary>
        /// <param name="ipValue"></param>
        /// <param name="outlet"></param>
//-----------------------------------------------snmpHalt------------------------------------------------------------------------
        private void snmpHalt(string ipValue, int outlet)
        {
            // Prepare target
            UdpTarget target = new UdpTarget((IPAddress)new IpAddress(ipValue));
            // Create a SET PDU
            Pdu pdu = new Pdu(PduType.Set);
            pdu.VbList.Add(new Oid("1.3.6.1.4.1.850.100.1.10.2.1.4." + outlet.ToString()), new Integer32(1));
            AgentParameters aparam = new AgentParameters(SnmpVersion.Ver2, new OctetString("tripplite"));
            // Response packet
            SnmpV2Packet response;
            try
            {
                // Send request and wait for response
                response = target.Request(pdu, aparam) as SnmpV2Packet;
            }
            catch (Exception ex)
            {
                // If exception happens, it will be returned here
                target.Close();
                return;
            }
            // Make sure we received a response
            if (response == null)
            {

            }
            else
            {
                // Check if we received an SNMP error from the agent
                if (response.Pdu.ErrorStatus != 0)
                {

                }
                else
                {
                    // Everything is ok. Agent will return the new value for the OID we changed

                }
            }
        }
        /// <summary>
        /// The function for "Halting" the racks.
        /// Double click the label which has the name of the rack appear on the right side. 
        /// Turns off every device in the racks utilizing a for loop 1-24.
        /// It also uses form5 which displays a window that states: "Please wait..."
        /// Prevents halt on rack 1 due to the internet patch...DO NOT CHANGE THAT FUNCTIONALITY. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
//----------------------------------------------------Double Click Label--------------------------------------------------------------
        private void textBox3_DoubleClick(object sender, EventArgs e)
        {
            string[] rac = treeView1.SelectedNode.Name.Split(' ');
            if (rac[0] == "R")
            {
                if (rac[1] == "1")
                {
                    //DO NOT CHANGE THIS IF STATEMENT, PROTECTS THE INTERNET PATCH
                    MessageBox.Show("Cannot shut down rack because of the Internet Switches.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    //Make a pop up option box to make sure that they want to shut down the entire rack. 
                    DialogResult dialogResult = MessageBox.Show("Are you sure you would like to shut down the entire rack?", "Rack Shut Down", MessageBoxButtons.YesNo);
                    //If they continue and shut down rack enter this if statement.
                    if (dialogResult == DialogResult.Yes)
                    {
                        int ip;
                        //fill in the ip and change from string to int
                        int.TryParse(rac[1], out ip);
                        ip += 89;
                        string ipnum = "132.177.122." + ip.ToString();
                        //show the please wait screen
                        PleaseWait loadingBar = new PleaseWait();
                        loadingBar.Show();
                        for (int i = 1; i < 25; i++)
                        {
                            //call snmpHalt for all outlets in the selected racks. 
                            //Turns off all outlets in the racks. 
                            snmpHalt(ipnum, i);
                            Thread.Sleep(2000);
                        }
                        //close the please wait screen.
                        loadingBar.Close();
                        //need to make the colors of the devices update themselves.
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        //Do nothing 
                    }


                }
            }
        }
        /// <summary>
        /// update field funtion updates the information on the right hand side of the screen.
        /// It is called in multiple locations to make sure the information is updating quickly when you select a new device. 
        /// </summary>
//------------------------------------------------------------Update Field-------------------------------------------------------------
        private void updateField()
        {
            //figure out what node is selected
            TreeNode node = treeView1.SelectedNode;
            //as long as the node is initialized
            if (node != null)
            {
                //get the name of the node
                string name = node.Text;
                int index = treeView1.SelectedNode.Index + 1;
                //for each of the nodes in the treeview
                for (int z = 0; z < strArray.Length - 1; z++)
                {
                    //in the case where name contains rack fill in information for general rack
                    if (name.Contains("Rack"))
                    {
                        textBox3.Text = treeView1.SelectedNode.Text;
                        textBox3.ForeColor = Color.Blue;
                        textBox4.Text = "None";
                        //set text to "on" or off depending on the color of the text
                        if (treeView1.SelectedNode.ForeColor == Color.LimeGreen)
                        {
                            label3.Text = "Device Power Status: On";

                        }
                        else if (treeView1.SelectedNode.ForeColor == Color.Red)
                        {
                            label3.Text = "Device Power Status: Off";

                        }
                        //if the color is black make the status of the device unknown since that means that it is not connecting to the GUI
                        else
                        {
                            label3.Text = "Device Power Status: Unknown";
                        }
                        textBox2.Text = " ";
                    }
                    //make the case for if the outlet is an "Empty" outlet
                    else if (name.Contains("Empty"))
                    {
                        textBox3.Text = "Empty Slot";
                        textBox3.ForeColor = Color.Black;
                        textBox4.Text = "None";
                        if (treeView1.SelectedNode.ForeColor == Color.LimeGreen && label3.Text != "Device Power Status: Powering Off...")
                        {
                            label3.Text = "Device Power Status: On";

                        }
                        else if (treeView1.SelectedNode.ForeColor == Color.Red && label3.Text != "Device Power Status: Powering On...")
                        {
                            label3.Text = "Device Power Status: Off";

                        }
                        else
                        {
                            label3.Text = "Device Power Status: Unknown";
                        }
                        textBox2.Text = " ";
                    }
                    //if it actually is an occupied outlet fill in the information on the device accordingly
                    else if (name == (index.ToString() + ". " + strArray[z][0]) && strArray[z][0] != "")
                    {
                        string[] oCompare = strArray[z][4].Split('&');
                        for( int i = 0; i <  oCompare.Length; i++)
                        {
                            if(index.ToString() == oCompare[i])
                            {
                                textBox3.Text = strArray[z][0];
                                textBox3.ForeColor = Color.Black;
                                textBox4.Text = strArray[z][1];
                                int racnum;
                                int outnum;
                                //get the outlet and rack numbers to fill in
                                string rac = strArray[z][3];
                                string outlet = strArray[z][4];
                                int.TryParse(rac, out racnum);
                                int.TryParse(outlet, out outnum);
                                int bip;
                                int.TryParse(strArray[z][3], out bip);
                                bip += 89;
                                int ou;
                                string ol = strArray[z][4];
                                int.TryParse(strArray[z][4], out ou);
                                //makes sure that the values that are being read follow through to all outlets that the device is hooked up to
                                string[] ands = ol.Split('&');
                                if (ands.Length > 1)
                                {
                                    //fill in information for a device with more than one outlet. 
                                    for (int h = 0; h < ands.Length; h++)
                                    {
                                        int o;
                                        bool isOutlet = int.TryParse(ands[h], out o);
                                        snmpGet(o, "132.177.122." + bip.ToString());
                                    }
                                }
                                else
                                {
                                    //check the status of the device
                                    snmpGet(ou, "132.177.122." + bip.ToString());
                                }
                                if (treeView1.SelectedNode.ForeColor == Color.LimeGreen && label3.Text != "Device Power Status: Powering Off...")
                                {
                                    //if the text is green write the status as being on to text box
                                    label3.Text = "Device Power Status: On";

                                }
                                else if (treeView1.SelectedNode.ForeColor == Color.Red && label3.Text != "Device Power Status: Powering On...")
                                {
                                    //write the status as being off to text box
                                    label3.Text = "Device Power Status: Off";

                                }
                                else
                                {
                                    //if the outlet is unable to connect fill in unkown status.
                                    label3.Text = "Device Power Status: Unknown";
                                }
                                textBox2.Text = " " + strArray[z][2];
                                return;
                            }
                        }
                    }
                }
            }            
        }
        /// <summary>
        /// timer tick is a function that runs every ten seconds. 
        /// It looks through the currently expanded racks and checks to see if there are any changes that the program missed. 
        /// Causes the program to lag slightly but overall improves the functionality of the program. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer1_Tick(object sender, EventArgs e)
        {
            //a for loop that represents all of the racks ( There are currently 16 racks )
            for (int j = 0; j < 17; j++)
            {
                //check to see what rack nodes in the treeview are currently expanded
                if (treeView1.Nodes[j].IsExpanded)
                {
                    //in the case that the rack node is expanded cycle through all of the outlets in that rack
                    for (int i = 0; i < 24; i++)
                    {
                        //get the ip if the outlet is currently connected, this method will not check any nodes in the tree that are currently disconnected
                        //due to issues that develop since it takes to long to check those values that the clock will make them check again by the time that they are done. 
                        if (treeView1.Nodes[j].Nodes[i].ForeColor == Color.Red || treeView1.Nodes[j].Nodes[i].ForeColor == Color.LimeGreen)
                        {
                            int ip = j + 90;
                            string ipNum = "132.177.122." + ip.ToString();
                            snmpGet(i + 1, ipNum);
                        }
                        else
                        {
                            //if one of the outlets in the tree is disconnected from the gui there is an error and the gui will automatically collapse the tree to 
                            //prevent further errors from occuring
                            treeView1.Nodes[j].Collapse();
                            return;
                        }
                    }
                }
            }
            //update the field for the current node. 
            updateField();
        }

        /// <summary>
        /// the mouseclick makes the selector change where it is hovering in the treeview object.
        /// This allows for us to select a new node for the fields on the right hand side of the 
        /// screen to update using.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            switch (e.Button)
            {
                case MouseButtons.Right:
                    if (e.Node.Level == 0)
                    {
                        treeView1.SelectedNode = treeView1.Nodes[e.Node.Index];
                    }
                    break;
            }

        }

        /// <summary>
        /// This creates a drop down menu that appears when one right clicks a node.
        /// This function will allow for racks to be updated and checked if they are connected mannually
        /// by selecting the ping option int the stripmenu. 
        /// When the option is selected the program will look at the highlighted rack and ping all of the 
        /// outlets in the rack to check if they are correctly connected. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void pingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //get the ipNum of the rack
            int ipNum = Int32.Parse(Regex.Match(treeView1.SelectedNode.Text, @"\d+").Value) + 89;
            string ips = "132.177.122." + ipNum.ToString();
            //check if the rack is connected
            bool isConnected = PingHost(ips);
            //debugging output to command line
            Console.WriteLine("HELLO");
            Console.WriteLine(isConnected);
            System.Diagnostics.Debug.WriteLine("Pinging " + ips);
            //If connected need to update the GUI and all of that racks child nodes ( Colors )
            if (isConnected == true)
            {
                try
                {
                    Dictionary<Tuple<int, int>, string> map = new Dictionary<Tuple<int, int>, string>();
                    //Boolean thingsOn = false;
                    for (int p = 1; p < 25; p++)
                    {
                        //change the color of the rack to correct color scheme
                        Tuple<int, int> tup = new Tuple<int, int>(ipNum, p);
                        map.Add(tup, ips);
                        if (treeView1.Nodes[ipNum - 90].Nodes[p - 1].ForeColor == Color.LimeGreen)
                        {
                            racTrue = true;
                        }
                    }

                    doWork(map);
                }
                catch (SnmpSharpNet.SnmpException me)
                {
                    //in the case that it has an error
                    Console.WriteLine("There was and error" + me);
                }
            }
            else
            {
                //in the case that the rack is still not connecting correctly sets the color to black
                for (int p = 1; p < 25; p++)
                {
                    //debugging output
                    Console.WriteLine("YAY");
                    treeView1.Nodes[ipNum - 90].Nodes[p - 1].ForeColor = Color.Black;
                    racTrue = false;
                }
            }
        }

        private void onSwitch_Click(object sender, EventArgs e)
        {
            //controls the toggle to turn on/off the outlets in the rack. 
            if (treeView1.SelectedNode != null & treeView1.SelectedNode.Text.Contains("Rack"))
            {

            }
            else
            {
                int ips = treeView1.SelectedNode.Parent.Index;
                int ol = treeView1.SelectedNode.Index;
                string ip = "132.177.122.";
                ips = ips + 90;
                ol = ol + 1;
                ip = ip + ips;
                //THIS PROTECTS THE INTERNET PATCH IN RACK1 FROM BEING TURNED OFF DO NOT CHANGE UNLESS YOU NEED TO CHANGE THE IP OF ALL THE RACKS
                if (ip == "132.177.122.90" && (ol == 18 || ol == 19 || ol == 20 || ol == 21 || ol == 22 || ol == 23 || ol == 24))
                {
                    MessageBox.Show("Power control for internet switches not allowed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    //calls the function that actually toggles the power in the rack. 
                    snmpSet(ip, ol, 2);
                }
            }
        }

        private void offSwitch_Click(object sender, EventArgs e)
        {
            //controls the toggle to turn on/off the outlets in the rack. 
            if (treeView1.SelectedNode != null & treeView1.SelectedNode.Text.Contains("Rack"))
            {

            }
            else
            {
                int ips = treeView1.SelectedNode.Parent.Index;
                int ol = treeView1.SelectedNode.Index;
                string ip = "132.177.122.";
                ips = ips + 90;
                ol = ol + 1;
                ip = ip + ips;
                //THIS PROTECTS THE INTERNET PATCH IN RACK1 FROM BEING TURNED OFF DO NOT CHANGE UNLESS YOU NEED TO CHANGE THE IP OF ALL THE RACKS
                if (ip == "132.177.122.90" && (ol == 18 || ol == 19 || ol == 20 || ol == 21 || ol == 22 || ol == 23 || ol == 24))
                {
                    MessageBox.Show("Power control for internet switches not allowed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    //calls the function that actually toggles the power in the rack. 
                    snmpSet(ip, ol, 1);
                }
            }
        }

        private void editToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (treeView1.SelectedNode.Parent == null || treeView1.SelectedNode.Text.Contains("Internet Patch"))
            {
                MessageBox.Show("You cannot make changes to the Rack, please select a device before proceeding.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                timer1.Stop();
                Cursor = Cursors.WaitCursor;
                var n = treeView1.SelectedNode;
                treeView1.Nodes[treeView1.SelectedNode.Parent.Index].Nodes[treeView1.SelectedNode.Index].Text = ((treeView1.SelectedNode.Index + 1).ToString() + ". Empty");
                EditForm editForm = new EditForm(strArray, textBox3.Text, treeView1.SelectedNode.Index + 1);
                editForm.ShowDialog();
                strArray = editForm.updateArray;
                Cursor = Cursors.WaitCursor;
                for (int i = 0; i < strArray.Length; i++)
                {
                    if (strArray[i][0] == "")
                    {
                        
                    }
                    else
                    {
                        string rac = strArray[i][3];
                        string ol = strArray[i][4];
                        //fills in the names on the treeview1 object.
                        //splits the values of the outlets at the & symbol
                        string[] ands = ol.Split('&');
                        if (ands.Length > 1)
                        {
                            //meant for things that hvae more than 2 outlets
                            for (int h = 0; h < ands.Length; h++)
                            {
                                //convert each of the string in the array from string to int to fill in the treeview object correctly
                                int r;
                                int o;
                                bool isOutlet = int.TryParse(ands[h], out o);
                                bool isRack = int.TryParse(rac, out r);
                                //decides the location of the treeview node to fill in and what to fill in there.
                                if (isRack && isOutlet)
                                    treeView1.Nodes[r - 1].Nodes[o - 1].Text = o + ". " + strArray[i][0];
                            }
                        }
                        //meant for things with only one outlet
                        else
                        {
                            //r = rack number
                            int r;
                            //o = outlet number
                            int o;
                            bool isOutlet = int.TryParse(ol, out o);
                            bool isRack = int.TryParse(rac, out r);
                            //decides the location of the treeview node to fill in and what to fill in there.
                            if (isRack && isOutlet)
                                treeView1.Nodes[r - 1].Nodes[o - 1].Text = o + ". " + strArray[i][0];
                        }
                    }
                }
                treeView1.SelectedNode = n;
                updateField();
                timer1.Start();
                Cursor = Cursors.Arrow;
            }
        }
        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (treeView1.SelectedNode.Parent == null || treeView1.SelectedNode.Text.Contains("Internet Patch"))
            {
                MessageBox.Show("You cannot delete the entire Rack, please select a device before proceeding.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                DialogResult dialogResult = MessageBox.Show("Are you sure that you would like to delete this device information? By deleteing this device info" + 
                    " it will be perminately deleted from the excel document. Continue?", "Delete", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    var n = treeView1.SelectedNode;
                    timer1.Stop();
                    Cursor.Current = Cursors.WaitCursor;
                    uint excelId = 0;
                    Excel.Application ExcelObj = new Excel.Application();
                    if (ExcelObj == null)
                    {
                        MessageBox.Show("Unable to connect to Microsoft Excel! Terminating program.", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        System.Windows.Forms.Application.Exit();
                    }
                    Excel.Sheets sheets;
                    Excel.Worksheet worksheet;
                    //opens the excel book and selects the first sheet
                    var books = ExcelObj.Workbooks;
                    var theWorkbook = books.Open(@"\\fox\raid\ethernets\Interop\Interop Randomizer\List of Link Partners.xls", 0, false, 5,
                            "", "", true, Excel.XlPlatform.xlWindows, "\t", true, false,
                            0, true);
                    //gets all the sheets in the workbook
                    sheets = theWorkbook.Worksheets;
                    int arrRow = -1;
                    GetWindowThreadProcessId(new IntPtr(ExcelObj.Hwnd), out excelId);
                    for (int i = 0; i < strArray.Length - 1; i++)
                    {
                        if (textBox3.Text == strArray[i][0] && textBox4.Text == strArray[i][1])
                        {
                            string ol = strArray[i][4];
                            //fills in the names on the treeview1 object.
                            //splits the values of the outlets at the & symbol
                            bool containsOutlet = false;
                            string[] andz = ol.Split('&');
                            if (andz.Length > 1)
                            {
                                //meant for things that hvae more than 2 outlets
                                for (int h = 0; h < andz.Length; h++)
                                {
                                    //convert each of the string in the array from string to int to fill in the treeview object correctly
                                    int o;
                                    bool isOutlet = int.TryParse(andz[h], out o);
                                    if (o == n.Index + 1)
                                        containsOutlet = true;
                                }
                            }
                            else
                            {
                                int o;
                                bool isOutlet = int.TryParse(strArray[i][4], out o);
                                if (o == n.Index + 1)
                                    containsOutlet = true;
                            }
                            if (containsOutlet == true)
                            {
                                arrRow = i;
                            }
                        }
                    }
                    
                    int sheetNum = -1;
                    int rowNum = -1;
                    Int32.TryParse(strArray[arrRow][strArray[arrRow].Length - 2], out sheetNum);
                    Int32.TryParse(strArray[arrRow][strArray[arrRow].Length - 1], out rowNum);
                    string[] outlets = strArray[arrRow][4].Split('&');
                    string rack = strArray[arrRow][3];
                    string[] temp = new string[1];
                    temp[0] = "";
                    strArray[arrRow] = temp;
                    worksheet = (Excel.Worksheet)sheets.get_Item(sheetNum);
                        ((Excel.Range)worksheet.Rows[rowNum]).Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                        theWorkbook.Save();
                        //close all excel objects that are currently open
                        theWorkbook.Close(false);
                        ExcelObj.Quit();
                        CloseExcel(sheets);
                        CloseExcel(theWorkbook);
                        CloseExcel(ExcelObj);
                        //the following try catch and process tracking was a solution from Jordy "Kaiwa" Ruiter found on
                        //www.codeproject.com/Questions/74980/Close-Excel-Process-with-Interop
                        try
                        {
                            if (excelId != 0)
                            {
                                Process excel = Process.GetProcessById((int)excelId);
                                excel.CloseMainWindow();
                                excel.Refresh();
                                excel.Kill();
                            }
                        }
                        catch
                        {
                            //process was already killed
                        }
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    for( int i = 0; i < outlets.Length; i ++ )
                    {
                        int r;
                        int o;
                        Int32.TryParse(rack, out r);
                        Int32.TryParse(outlets[i], out o);
                        treeView1.Nodes[r - 1].Nodes[o - 1].Text = o.ToString() + ". Empty";
                    }
                    for( int j = 0; j < strArray.Length; j ++ )
                    {
                        if( strArray[j][0] != "" )
                        {
                            int temp1;
                            Int32.TryParse(strArray[j][strArray[arrRow].Length - 1], out temp1);
                            if (temp1 > rowNum)
                            {
                                temp1 = temp1 - 1;
                                strArray[j][strArray[arrRow].Length - 1] = temp1.ToString();
                            }
                        }
                    }
                    treeView1.SelectedNode = n;
                    updateField();
                    timer1.Start();
                    Cursor.Current = Cursors.WaitCursor;
                }
            }
        }

        private void refreshSwitch_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This Feature Has Not Been Implemented Yet", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            /*
            timer1.Stop();
            Cursor.Current = Cursors.WaitCursor;
            treeView1.CollapseAll();
            for (int r = 0; r < 17; r++)
            {
                for (int o = 0; o < 24; o++)
                {
                    string[] s = new string[1];
                    s[0] = "";
                    strArray[((r + 1) * (o + 1)) - 1] = s;
                    treeView1.Nodes[r].Nodes[o].Text = (o + 1).ToString() + ". Empty";
                }
            }
            uint excelId = 0;
            Excel.Application ExcelProc = new Excel.Application();
            // CHECK IF EXCEL IS AVAIABLE,
            // if not, close the program as it needs excel.
            if (ExcelProc == null)
            {
                MessageBox.Show("Unable to connect to Microsoft Excel! Terminating program.", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Windows.Forms.Application.Exit();
            }
            Excel.Sheets sheets;
            Excel.Worksheet worksheet;
            //opens the excel book and selects the first sheet
            var books = ExcelProc.Workbooks;
            //opens the excel book and selects the first sheet
            var book = ExcelProc.Workbooks;
            var theWorkbook = book.Open(
                @"\\fox\raid\ethernets\Interop\Interop Randomizer\List of Link Partners.xls", 0, true, 5,
                    "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false,
                    0, true);
            //gets all the sheets in the workbook
            sheets = theWorkbook.Worksheets;
            int l = 1;
            int loc = 0;
            GetWindowThreadProcessId(new IntPtr(ExcelProc.Hwnd), out excelId);
            //traverses through the excel worksheets until there are no more pages
            while (l <= sheets.Count)
            {
                worksheet = (Excel.Worksheet)sheets.get_Item(l);
                int workSheetCount = sheets.Count;
                int i = 2;
                //creates the range of cells to look at
                var myRange = worksheet.get_Range("A" + i);
                int count = 0;
                //uses the myRange value to know when it is at the end of the spreadsheet page. 
                while (myRange.Value2 != null)
                {
                    Excel.Range iolRange = worksheet.get_Range("E" + i);
                    //if-else statement that figures out if the device is part of the racks and calculates the last 
                    // location to be filled. Makes sure there is no empty space in the array list.
                    if (iolRange.Value2 == null)
                    {
                        //if the cell is blank it does this
                        count += 1;
                        i += 1;
                        myRange = worksheet.get_Range("A" + i);
                    }
                    else
                    {
                        //if the cell has a value it performs these functions to insert the values into the correct location
                        //in the rack arrays/treeview object
                        myRange = worksheet.get_Range("A" + i);
                        var range = worksheet.get_Range("A" + i, "F" + i);
                        System.Array myvalues = (System.Array)range.Cells.Value;
                        strArray[i - (2 + count) + loc] = ConvertToStringArray(myvalues);
                        if (strArray[i - (2 + count) + loc][0] != "")
                        {
                            //stores the name of the worksheet
                            strArray[i - (2 + count) + loc][strArray[0].Length - 3] = worksheet.Name;
                            //stores the value of the worksheet the item was found on so that we can make edits later
                            strArray[i - (2 + count) + loc][strArray[0].Length - 2] = (string)l.ToString();
                            //stores the value of the original row of the spreadsheet the item was found
                            strArray[i - (2 + count) + loc][strArray[0].Length - 1] = (string)i.ToString();
                        }
                        string rac = strArray[i - (2 + count) + loc][3];
                        string ol = strArray[i - (2 + count) + loc][4];
                        //fills in the names on the treeview1 object.
                        //splits the values of the outlets at the & symbol
                        string[] ands = ol.Split('&');
                        if (ands.Length > 1)
                        {
                            //meant for things that hvae more than 2 outlets
                            for (int h = 0; h < ands.Length; h++)
                            {
                                //convert each of the string in the array from string to int to fill in the treeview object correctly
                                int r;
                                int o;
                                bool isOutlet = int.TryParse(ands[h], out o);
                                bool isRack = int.TryParse(rac, out r);
                                //decides the location of the treeview node to fill in and what to fill in there.
                                if (isRack && isOutlet)
                                    treeView1.Nodes[r - 1].Nodes[o - 1].Text = o + ". " + strArray[i - (2 + count) + loc][0];
                            }
                        }
                        //meant for things with only one outlet
                        else
                        {
                            //r = rack number
                            int r;
                            //o = outlet number
                            int o;
                            bool isOutlet = int.TryParse(ol, out o);
                            bool isRack = int.TryParse(rac, out r);
                            //decides the location of the treeview node to fill in and what to fill in there.
                            if (isRack && isOutlet)
                                treeView1.Nodes[r - 1].Nodes[o - 1].Text = o + ". " + strArray[i - (2 + count) + loc][0];
                        }
                        i += 1;
                        myRange = worksheet.get_Range("A" + i);
                        CloseExcel(range);
                    }
                    CloseExcel(iolRange);
                }
                loc = loc + i - (2 + count);
                l += 1;
                CloseExcel(myRange);
            }
            //close all excel objects that are currently open
            theWorkbook.Close(false);
            ExcelProc.Quit();
            CloseExcel(sheets);
            CloseExcel(book);
            CloseExcel(theWorkbook);
            CloseExcel(ExcelProc);
            //the following try catch and process tracking was a solution from Jordy "Kaiwa" Ruiter found on
            //www.codeproject.com/Questions/74980/Close-Excel-Process-with-Interop
            try
            {
                if (excelId != 0)
                {
                    Process excel = Process.GetProcessById((int)excelId);
                    excel.CloseMainWindow();
                    excel.Refresh();
                    excel.Kill();
                }
            }
            catch
            {
                //process was already killed
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            timer1.Start();
            Cursor.Current = Cursors.Arrow;
            */
        } 
    } 
}