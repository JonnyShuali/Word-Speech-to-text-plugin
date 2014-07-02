using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Word;
using System.Collections;

namespace wordSpeech
{
    public partial class Ribbon1
    {
        read x;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            refreshDevices();
        }
        public void refreshDevices()
        {
    ArrayList devices=microphone.DeviceList();
            foreach( RibbonDropDownItem item in devices)
            {
                micChooser.Items.Add(item);
            }
        }

        private void reader_Click(object sender, RibbonControlEventArgs e)
        {
            if(reader.Checked.Equals(true))
            x = new read(wordHelper.getDoc());
            else
            {
            if (x != null)
                x.stopReading();
        }
        }
                 public  void readEnd(object sender, EventArgs e)
        {
            reader.Checked = false;
        }
        

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if (x != null)
                x.stopReading();
        }



    }
}
