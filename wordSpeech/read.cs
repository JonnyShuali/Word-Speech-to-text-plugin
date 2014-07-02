using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Speech.Synthesis;
using System.Speech.Recognition;

namespace wordSpeech
{
    class read
    {
        SpeechSynthesizer sSynth;
        PromptBuilder pBuilder;
        public read()
        {
            sSynth = new SpeechSynthesizer();
            pBuilder = new PromptBuilder();

        }
          public read(String text)
        {
            sSynth = new SpeechSynthesizer();
            pBuilder = new PromptBuilder();
            readText(text);

        }
        public  void readText(String text)
        {
            pBuilder.ClearContent();
            pBuilder.AppendText(text);
            //sSynth.Speak(pBuilder);
            sSynth.SpeakAsync(pBuilder);
            sSynth.SpeakCompleted += (s,e)=>
            Globals.Ribbons.Ribbon1.reader.Checked = false;
        }

        public void stopReading()
        {
            //sSynth.SpeakAsyncCancel(pBuilder);
          sSynth.Pause();
        }
    }
}
