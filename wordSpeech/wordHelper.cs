using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;

namespace wordSpeech
{
     class wordHelper
    {
        public static void insertText(String input)
        {
            Application wordApp = (Microsoft.Office.Interop.Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
            Microsoft.Office.Interop.Word.Selection currentSelection = wordApp.Selection;

            // Store the user's current Overtype selection 
            bool userOvertype = wordApp.Options.Overtype;

            // Make sure Overtype is turned off. 
            if (wordApp.Options.Overtype)
            {
                wordApp.Options.Overtype = false;
            }

            // Test to see if selection is an insertion point. 
            if (currentSelection.Type == Microsoft.Office.Interop.Word.WdSelectionType.wdSelectionIP)
            {
                currentSelection.TypeText(input);
                currentSelection.TypeParagraph();
            }
            else
                if (currentSelection.Type == Microsoft.Office.Interop.Word.WdSelectionType.wdSelectionNormal)
                {
                    // Move to start of selection. 
                    if (wordApp.Options.ReplaceSelection)
                    {
                        object direction = Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseStart;
                        currentSelection.Collapse(ref direction);
                    }
                    currentSelection.TypeText("Inserting before a text block. ");
                    currentSelection.TypeParagraph();
                }
                else
                {
                    // Do nothing.
                }

            // Restore the user's Overtype selection
            wordApp.Options.Overtype = userOvertype;
        }
      public static String getDoc()
        {
            Application wordApp = (Microsoft.Office.Interop.Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
            Microsoft.Office.Tools.Word.Document vstoDocument = Globals.Factory.GetVstoObject(wordApp.Application.ActiveDocument);
            return vstoDocument.Content.Text;

        }
         public static String getLang()
         {
                         Application wordApp = (Microsoft.Office.Interop.Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
            Microsoft.Office.Interop.Word.Selection currentSelection = wordApp.Selection;
            Microsoft.Office.Tools.Word.Document vstoDocument = Globals.Factory.GetVstoObject(wordApp.Application.ActiveDocument);
            vstoDocument.DetectLanguage();
            if (vstoDocument.LanguageDetected)
                return "English";
            return "elese";
         }
        }
    }

