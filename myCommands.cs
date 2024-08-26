// (C) Copyright 2024 by www.lisp.vn
//
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(AJS_RibbonAddings.MyCommands))]

namespace AJS_RibbonAddings
{
    // This class is instantiated by AutoCAD for each document when
    // a command is called by the user the first time in the context
    // of a given document. In other words, non static data in this class
    // is implicitly per-document!
    public class MyCommands
    {
        // Modal Command with localized name
        [CommandMethod("AJSRibbonAddings", CommandFlags.Modal)]
        public void MyCommand() // This method can have any name
        {
            // Put your command code here
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed;
            if (doc != null)
            {
                ed = doc.Editor;
                ed.WriteMessage("Hello, this is your first command.");
                MyPlugin.ToggleButton();
            }
        }
    }
}