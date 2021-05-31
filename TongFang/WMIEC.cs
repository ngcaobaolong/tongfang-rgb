using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace TongFang
{

    public static class WMIEC
    {
        static public String WMIWriteECRAM(UInt64 Addr, UInt64 Value)
        {
            try
            {
                //wbemtest
                ManagementObject classInstance =
                    new ManagementObject("root\\WMI",
                    "AcpiTest_MULong.InstanceName='ACPI\\PNP0C14\\1_1'",
                    null);

                // Obtain in-parameters for the method
                ManagementBaseObject inParams =
                    classInstance.GetMethodParameters("GetSetULong");

                // Add the input parameters.
                Value <<= 16;
                Addr = 0x0000000000000000 + Value + Addr;
                inParams["Data"] = Addr;

                // Execute the method and obtain the return values.
                // Wordaround to avoid busy flag
                //System.Threading.Thread.Sleep(200);
                ManagementBaseObject outParams =
                    classInstance.InvokeMethod("GetSetULong", inParams, null);
                // List outParams
                return outParams["Return"].ToString();

            }
            catch (ManagementException err)
            {
                Console.WriteLine("GetSetULong failed: " + err.Message);
                return "Failed";
            }
        }
        static public String PermaECRAM(UInt64 Addr, UInt64 Value)
        {
            try
            {
                //wbemtest
                ManagementObject classInstance =
                    new ManagementObject("root\\WMI",
                    "AcpiODM_Demo.InstanceName='ACPI\\PNP0C14\\1_1'",
                    null);

                // Obtain in-parameters for the method
                ManagementBaseObject inParams =
                    classInstance.GetMethodParameters("GetUlongEx7");

                // Add the input parameters.
                Value <<= 16;
                Addr = 0x0000000000000000 + Value + Addr;
                inParams["Data"] = Addr;

                // Execute the method and obtain the return values.
                // Wordaround to avoid busy flag
                //System.Threading.Thread.Sleep(200);
                ManagementBaseObject outParams =
                    classInstance.InvokeMethod("GetUlongEx7", inParams, null);
                // List outParams
                return outParams["Return"].ToString();

            }
            catch (ManagementException err)
            {
                Console.WriteLine("GetUlongEx7 failed: " + err.Message);
                return "Failed";
            }
        }
    }
}
