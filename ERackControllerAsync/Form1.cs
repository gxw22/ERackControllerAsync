using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OpcRcw.Comn;
using OpcRcw.Da;
using System.Runtime.InteropServices;

namespace ERackControllerAsync
{
    public partial class Form1 : Form
    {
        private const string SERVER_NAME = "OPC.SimaticNET";
        internal const string ITEM1_NAME = "S7:[S7 connection_1]DB10,INT0";          // 1st item name
        internal const string ITEM2_NAME = "S7:[S7 connection_1]MX200.0";          // 2nd item name
        internal const string GROUP_NAME = "opc_ethernet";
        internal const int LOCALE_ID = 0x409;                       // LOCALE FOR ENGLISH.

        /* Global variables */
        IOPCServer pIOPCServer;
        IOPCAsyncIO2 pIOPCAsyncIO2 = null;                          // instance pointer for asynchronous IO.
        IOPCGroupStateMgt pIOPCGroupStateMgt = null;
        IConnectionPointContainer pIConnectionPointContainer = null;
        IConnectionPoint pIConnectionPoint = null;

        Object pobjGroup1 = null;
        int pSvrGroupHandle = 0;                                    // server group handle for the added group
        int nTransactionID = 0;
        int[] ItemSvrHandleArray;
        Int32 dwCookie = 0;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            // Local variables
            Type svrComponenttyp;
            OPCITEMDEF[] ItemDeffArray;

            // Group properties
            Int32 dwRequestedUpdateRate = 250;
            Int32 hClientGroup = 1;
            Int32 pRevUpdateRate;
            float deadband = 0;
            int TimeBias = 0;
            GCHandle hTimeBias, hDeadband;
            hTimeBias = GCHandle.Alloc(TimeBias, GCHandleType.Pinned);
            hDeadband = GCHandle.Alloc(deadband, GCHandleType.Pinned);

            // 1. Get the Type from the progID and create instance of the OPC Server COM component
            Guid iidRequiredInterface = typeof(IOPCItemMgt).GUID;
            svrComponenttyp = Type.GetTypeFromProgID(SERVER_NAME);
            try
            {
                // Connect to the local server.
                pIOPCServer = (IOPCServer)Activator.CreateInstance(svrComponenttyp);
            }
            catch
            { 
            
            }
        }

        private void AddGroup()
        {
            /* 2. Add a new group
                        Add a group object and querry for interface IOPCItemMgt
                        Parameter as following:
                        [in] not active, so no OnDataChange callback
                        [in] Request this Update Rate from Server
                        [in] Client Handle, not necessary in this sample
                        [in] No time interval to system UTC time
                        [in] No Deadband, so all data changes are reported
                        [in] Server uses english language to for text values
                        [out] Server handle to identify this group in later calls
                        [out] The answer from Server to the requested Update Rate
                        [in] requested interface type of the group object
                        [out] pointer to the requested interface
                    */

            pIOPCServer.AddGroup(GROUP_NAME,
                0,
                dwRequestedUpdateRate,
                hClientGroup,
                hTimeBias.AddrOfPinnedObject(),
                hDeadband.AddrOfPinnedObject(),
                LOCALE_ID,
                out pSvrGroupHandle,
                out pRevUpdateRate,
                ref iidRequiredInterface,
                out pobjGroup1);
        }
    }
}
