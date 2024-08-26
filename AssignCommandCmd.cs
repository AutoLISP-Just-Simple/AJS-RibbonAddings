using Autodesk.AutoCAD.ApplicationServices;
using System;

namespace AJS_RibbonAddings
{
    public class AssignCommandCmd : System.Windows.Input.ICommand
    {
        public string cmd = "";

        public AssignCommandCmd(string _cmd)
        {
            cmd = _cmd.Replace("_.", "").Trim();
            if (cmd.Left(1) == "_") cmd = cmd.Mid(1, 100);
            if (cmd.Left(1) == ".") cmd = cmd.Mid(1, 100);
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public event EventHandler CanExecuteChanged;

        public void Execute(object parameter)
        {
            Run(cmd);
        }

        private static void Run(string cmd)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (var dl = doc.LockDocument())
            {
                string esc = "";

                string cmds = (string)Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CMDNAMES");

                if (cmds.Length > 0)
                {
                    int cmdNum = cmds.Split(new char[] { '\'' }).Length;
                    for (int i = 0; i < cmdNum; i++)
                        esc += '\x03';
                }

                doc.SendStringToExecute(esc + "_." + cmd.Trim() + " ", true, false, true);
            }
        }
    }
}