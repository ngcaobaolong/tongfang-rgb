using System;
using System.Linq;
using System.Drawing;
using System.Collections.Generic;
using HidSharp;
using HidSharp.Reports;
using HidSharp.Reports.Encodings;
using System.Threading;

namespace TongFang
{
    public static class Keyboard
    {
        #region Constants
        private const int VID = 0x048D;
        private const int PID = 0xCE00;
        private const uint USAGE_PAGE = 0xFF03;
        private const uint USAGE = 0x001;
        private const byte ROWS = 6;
        private const byte COLUMNS = 21;
        #endregion

        #region Fields
        private static HidDevice _device;
        private static HidStream _deviceStream;
        private static readonly Color[] _colors = new Color[126];
        private static Dictionary<Key, byte> _layout;
        private static bool _dirty = true;
        #endregion

        #region Properties
        public static bool IsConnected { get; private set; }
        #endregion

        /// <summary>
        /// Tries to initialize a connection to the keyboard. 
        /// </summary>
        /// <param name="brightness">Brightness value, between 0 and 100. Defaults to 50.</param>
        /// <param name="layout">ISO or ANSI. Defaults to ANSI</param>
        /// <returns>Returns true if successful.</returns>
        public static bool Initialize(int brightness = 100, Layout lyt = Layout.ANSI)
        {
            var devices = DeviceList.Local.GetHidDevices(VID).Where(d => d.ProductID == PID);

            if (!devices.Any())
                return false;

            try
            {
                _device = GetFromUsages(devices, USAGE_PAGE, USAGE);

                if (_device?.TryOpen(out _deviceStream) ?? false)
                {
                    _layout = lyt == Layout.ANSI ? Layouts.ANSI : Layouts.ISO;
                    SetEffectType(Control.Default, Effect.UserMode, 0, (byte)(brightness / 2), 0, 0, 0);
                    return IsConnected = true;
                }
                else
                {
                    _deviceStream?.Close();
                }
            }
            catch
            { }

            return false;
        }

        /// <summary>
        /// Writes colors to the keyboard
        /// </summary>
        public static bool Update()
        {
            if (!_dirty)
                return true;
            //packet structure: 65 bytes
            //byte 0 = 0 ???
            //byte 1 = 0 ???
            //byte 2 to 22 = B
            //byte 23 to 43 = G
            //byte 44 to 64 = R

            var packet = new byte[65];
            try
            {
                for (byte row = 0; row < ROWS; row++)
                {
                    for (byte column = 0; column < COLUMNS; column++)
                    {
                        int colorIndex = column + ((5 - row) * 21);

                        packet[2 + column] = _colors[colorIndex].B;
                        packet[23 + column] = _colors[colorIndex].G;
                        packet[44 + column] = _colors[colorIndex].R;
                    }

                    SetRowIndex(row);
                    _deviceStream.Write(packet);
                    Thread.Sleep(1);
                }
            }
            catch
            {
                return false;
            }
            _dirty = false;
            return true;
        }

        /// <summary>
        /// Sets a given key to a given Color
        /// </summary>
        /// <param name="k">key to set</param>
        /// <param name="clr">color to set the key to</param>
        public static void SetKeyColor(Key k, Color clr)
        {
            if (!_layout.TryGetValue(k, out var idx))
                return;

            if (_colors[idx] != clr)
            {
                _colors[idx] = clr;
                _dirty = true;
            }
        }

        /// <summary>
        /// Sets every key to the same color
        /// </summary>
        /// <param name="clr"></param>
        public static void SetColorFull(Color clr)
        {
            for (int i = 0; i < _colors.Length; i++)
                _colors[i] = clr;
            
            _dirty = true;
        }

        /// <summary>
        /// Closes the connection to the keyboard
        /// </summary>
        public static void Disconnect()
        {
            _deviceStream?.Close();
            IsConnected = false;
        }

        #region Private methods
        private static bool SetEffectType(Control control, Effect effect, byte speed, byte light, byte colorIndex, byte direction, byte save)
        {
            byte[] buffer = new byte[9]
            {
                0,
                8,
                (byte)control,
                (byte)effect,
                speed,
                light,
                colorIndex,
                direction,
                save
            };
            try
            {
                _deviceStream.SetFeature(buffer);
            }
            catch
            {
                return false;
            }

            return true;
        }

        private static bool SetRowIndex(byte idx)
        {
            byte[] buffer = new byte[9]
            {
                0,
                22,
                0,
                idx,
                0,
                0,
                0,
                0,
                0
            };
            try
            {
                _deviceStream.SetFeature(buffer);
            }
            catch
            {
                return false;
            }

            return true;
        }

        private static HidDevice GetFromUsages(IEnumerable<HidDevice> devices, uint usagePage, uint usage)
        {
            foreach (var dev in devices)
            {
                try
                {
                    var raw = dev.GetRawReportDescriptor();
                    var usages = EncodedItem.DecodeItems(raw, 0, raw.Length).Where(t => t.TagForGlobal == GlobalItemTag.UsagePage);

                    if (usages.Any(g => g.ItemType == ItemType.Global && g.DataValue == usagePage))
                    {
                        if (usages.Any(l => l.ItemType == ItemType.Local && l.DataValue == usage))
                        {
                            return dev;
                        }
                    }
                }
                catch
                {
                    //failed to get the report descriptor, skip
                }
            }

            return null;
        }
        #endregion
    }
}
