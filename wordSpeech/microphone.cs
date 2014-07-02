using Microsoft.Office.Tools.Ribbon;
using NAudio.Wave;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace wordSpeech
{
    class microphone
    {
        public static ArrayList DeviceList()
        {
            int waveInDevices = WaveIn.DeviceCount;
            ArrayList arr = new ArrayList();
            for (int waveInDevice = 0; waveInDevice < waveInDevices; waveInDevice++)
            {
                WaveInCapabilities deviceInfo = WaveIn.GetCapabilities(waveInDevice);
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = String.Format("{0}, {1} channels",
                     deviceInfo.ProductName, deviceInfo.Channels);
                item.Tag = waveInDevice;
                arr.Add(item);
             
            }
            return arr;
        }
    }
}
