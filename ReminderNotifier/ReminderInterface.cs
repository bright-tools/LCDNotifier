using System;
using HidLibrary;
using System.Linq;
using System.Text;
using System.Threading;

public class ReminderInterface
{
    private const int VendorId = 0x1987;
    private const int ProductId = 0x1100;
    private const int ReportLength = 64;
    private const int DisplayStringLen = 40;

    private static HidDevice device;

    private static void DeviceAttachedHandler()
    {
    }

    private static void DeviceRemovedHandler()
    {
    }

    public static void Connect()
    {
        device = HidDevices.Enumerate(VendorId, ProductId).FirstOrDefault();
        if (device != null)
        {

            device.OpenDevice();
            device.Inserted += DeviceAttachedHandler;
            device.Removed += DeviceRemovedHandler;
            device.MonitorDeviceEvents = true;
            device.ReadReport(ReadReportCallback);
        }
    }

    protected static bool TryConnect()
    {
        bool retVal = false;

        if (device == null)
        {
            Connect();
        }

        if (device == null)
        {

        }
        else
        {
            retVal = true;
        }

        return retVal;
    }

    public static bool RemoveMessage( uint id )
    {
        bool retVal = false;

        if (TryConnect())
        {
            byte[] header = { 0, 0x20 };

            // We don't worry about endianism here - the actual value of the ID isn't important - it's just used
            //  to match previous messages.
            byte[] message = header.Concat(BitConverter.GetBytes(id)).ToArray();

            Array.Resize(ref message, ReportLength);

            if (device.WriteReport(new HidReport(ReportLength, new HidDeviceData(message, HidDeviceData.ReadStatus.Success))))
            {
                retVal = true;
            }
        }

        return retVal;
    }

    public static bool SendMessage(uint id, String line1, String line2)
    {
        bool retVal = false;

        if ( TryConnect() )
        {
            byte[] header = { 0, 0x10 };
            byte[] row = { 0 };
            byte[] padding = { 0, 0, 0, 0, 0, 0, 0, 0, 0, };

            byte[] message = header.Concat( BitConverter.GetBytes(id) ).
                                    Concat( row ).
                                    Concat( padding ).
                                    Concat(PackString(line1)).ToArray();

            Array.Resize(ref message, ReportLength);

            if (device.WriteReport(new HidReport(ReportLength, new HidDeviceData(message, HidDeviceData.ReadStatus.Success))))
            {
                retVal = true;
            }

            row[0] = 1;

            message = header.Concat(BitConverter.GetBytes(id)).
                                    Concat(row).
                                    Concat(padding).
                                    Concat(PackString(line2)).ToArray();

            Array.Resize(ref message, ReportLength);

            if (device.WriteReport(new HidReport(ReportLength, new HidDeviceData(message, HidDeviceData.ReadStatus.Success))))
            {
                retVal = true;
            }
        }

        return retVal;
    }

    private static byte[] PackString(string str1)
    {
        byte[] message = Encoding.ASCII.GetBytes(str1.PadRight(DisplayStringLen, (char)0));

        return message;
    }

    private static void ReadReportCallback(HidReport report)
    {
        Console.WriteLine("recv: {0}", string.Join(", ", report.Data.Select(b => b.ToString("X2"))));
        device.ReadReport(ReadReportCallback);
    }

    public ReminderInterface()
	{
        device = null;
    }

    ~ReminderInterface()
    {
        if (device != null)
        {
            device.CloseDevice();
        }
    }
}
