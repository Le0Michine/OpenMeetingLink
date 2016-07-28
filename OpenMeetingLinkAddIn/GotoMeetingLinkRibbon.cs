﻿using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new GotoMeetingLinkRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace OpenMeetingLinkAddIn
{
  [ComVisible(true)]
  public class GotoMeetingLinkRibbon : Office.IRibbonExtensibility
  {
    public void OpenMeetingLink(Office.IRibbonControl control)
    {
      var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
      if (explorer?.Selection != null && explorer.Selection.Count > 0)
      {
        AppointmentItem currentItem = explorer.Selection[1];
        Regex linkParser = new Regex(@"\b(?:https?://|www\.)\S+\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        foreach (Match match in linkParser.Matches(currentItem.Location))
        {
          System.Diagnostics.Process.Start(match.Value);
        }
      }
    }

    #region IRibbonExtensibility Members

    public string GetCustomUI(string ribbonID)
    {
      return GetResourceText("OpenMeetingLinkAddIn.GotoMeetingLinkRibbon.xml");
    }

    #endregion

    #region Ribbon Callbacks
    //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

    public void Ribbon_Load(Office.IRibbonUI ribbonUI)
    {
    }

    #endregion

    #region Helpers

    private static string GetResourceText(string resourceName)
    {
      var asm = Assembly.GetExecutingAssembly();
      var resourceNames = asm.GetManifestResourceNames();
      foreach (string t in resourceNames)
      {
        if (string.Compare(resourceName, t, StringComparison.OrdinalIgnoreCase) == 0)
        {
          // ReSharper disable once AssignNullToNotNullAttribute
          using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(t)))
          {
            return resourceReader.ReadToEnd();
          }
        }
      }
      return null;
    }

    #endregion
  }
}