
// (C) Copyright 2023 by
//
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Customization;
//using Autodesk.AutoCAD.Customization;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
using System.Diagnostics;
using System.IO;
using RibbonButton = Autodesk.Windows.RibbonButton;
using RibbonControl = Autodesk.Windows.RibbonControl;
using RibbonPanelSource = Autodesk.Windows.RibbonPanelSource;
using RibbonSplitButton = Autodesk.Windows.RibbonSplitButton;

// This line is not mandatory, but improves loading performances
[assembly: ExtensionApplication(typeof(AJS_RibbonAddings.MyPlugin))]

namespace AJS_RibbonAddings
{
    public class MyPlugin : IExtensionApplication
    {
        const string Caption = "AJS Ribbon Addings - www.lisp.vn";
        void IExtensionApplication.Initialize()
        {
            ComponentManager.ItemInitialized += (s, e) => ToggleButton();
        }

        void IExtensionApplication.Terminate()
        {
            // Do plug-in application clean up here
        }

        private static string AddInPath = typeof(MyPlugin).Assembly.Location;

        // Button icons directory
        private static string FolderPath = Path.GetDirectoryName(AddInPath);

        public static void ToggleButton()
        {
            if (F4 != null) return;

            RibbonControl ribbonControl = ComponentManager.Ribbon;

            if (ribbonControl == null) return;

            var Tab = ribbonControl.FindTab("ID_AJS_By_LISP_VN");
            if (Tab == null)
            {
                Tab = new RibbonTab();
                Tab.Title = "AutoLISP Just Simple - www.lisp.vn";
                Tab.Id = "ID_AJS_By_LISP_VN";

                ribbonControl.Tabs.Add(Tab);
            }

            if (Tab == null) Tab = ribbonControl.FindTab("ACAD.RBN_00012112");
            if (Tab != null)
            {
                var Panel = ribbonControl.FindPanel("ID_AJS_RibbonAddings_2024", false);
                if (Panel == null)
                {
                    Debug.Print("Found: Title=" + Tab.Title + " Id=" + Tab.Id + " UID=" + Tab.UID);
                    //Panel.AutomationName = "AutoCAD Debugger by AJS";

                    RibbonPanelSource srcPanel = new RibbonPanelSource();
                    srcPanel.Title = "Debuger";
                    //srcPanel.Items.Insert(2, new RibbonSeparator());

                    F4 = new RibbonButton();
                    F4.Id = "cmd_F4";
                    F4.Text = "Copy all";
                    //F4.Size = RibbonItemSize.Large;
                    //F4.Orientation = System.Windows.Controls.Orientation.Horizontal;
                    F4.ShowText = true;
                    F4.ShowImage = true;
                    F4.CommandParameter = "^C^C_F4";
                    F4.CommandHandler = new AssignCommandCmd("F4");
                    F4.Image = Properties.Resources.ID_ACETALIAS.ToImageSource();
                    F4.LargeImage = Properties.Resources.ID_ACETALIAS.ToImageSource();

                    F5 = new RibbonButton();
                    F5.Id = "cmd_F5";
                    F5.Text = "ProgressMeter";
                    //F5.Size = RibbonItemSize.Large;
                    //F5.Orientation = System.Windows.Controls.Orientation.Vertical;
                    F5.ShowText = true;
                    F5.ShowImage = true;
                    F5.CommandParameter = "^C^C_F5";
                    F5.CommandHandler = new AssignCommandCmd("F5");
                    F5.Image = Properties.Resources.ID_ACETALIAS.ToImageSource();
                    F5.LargeImage = Properties.Resources.ID_ACETALIAS.ToImageSource();

                    F6 = new RibbonButton();
                    F6.Id = "cmd_F6";
                    F6.Text = "COPYHIST";
                    //F6.Size = RibbonItemSize.Large;
                    //F6.Orientation = System.Windows.Controls.Orientation.Vertical;
                    F6.ShowText = true;
                    F6.ShowImage = true;
                    F6.CommandParameter = "^C^C_F6";
                    F6.CommandHandler = new AssignCommandCmd("F6");
                    F6.Image = Properties.Resources.ID_ACETALIAS.ToImageSource();
                    F6.LargeImage = Properties.Resources.ID_ACETALIAS.ToImageSource();

                    var pt = new RibbonSplitButton();
                    pt.Description = "Copy history and paste into Notepad++";
                    pt.Text = "Copy history";
                    pt.ShowText = true;

                    pt.Size = RibbonItemSize.Large;
                    //pt.Image = Properties.Resources.ID_ACETALIAS.ToImageSource();
                    //pt.LargeImage = Properties.Resources.ID_ACETALIAS.ToImageSource();

                    //pt.Orientation = System.Windows.Controls.Orientation.Vertical;

                    //pt.ShowImage = true;

                    pt.IsSplit = true;
                    pt.Size = RibbonItemSize.Large;
                    pt.IsSynchronizedWithCurrentItem = true;
                    pt.Items.Add(F4);
                    pt.Items.Add(F6);
                    pt.Items.Add(F6);
                    srcPanel.Items.Add(pt);

                    RibbonCombo rcb = new RibbonCombo();
                    rcb.Description = "Copy history and paste into Notepad++";
                    rcb.Text = "AJS System";
                    rcb.ShowText = true;

                    RibbonButton ribBut01 = new RibbonButton();
                    ribBut01.Text = "www.lisp.vn";
                    ribBut01.Id = "Item01";
                    ribBut01.ShowText = true;
                    //ribBut01.ShowImage = true;
                    //_rbNewSymbol.LargeImage = Img.getBitmap(Properties.Resources.New_Symbol_32);
                    //_rbNewSymbol.Orientation = System.Windows.Controls.Orientation.Vertical;

                    RibbonButton ribBut02 = new RibbonButton();
                    ribBut02.Text = "www.tankhanh.com.vn";
                    ribBut02.Id = "Item02";
                    ribBut02.ShowText = true;
                    //ribBut01.ShowImage = true;

                    rcb.Items.Add(ribBut01);
                    rcb.Items.Add(ribBut02);

                    srcPanel.Items.Add(rcb);

                    Panel = new RibbonPanel();
                    Panel.Id = "ID_AJS_RibbonAddings_2024";
                    Panel.Source = srcPanel;

                    Tab.Panels.Add(Panel);

                    bool assign_ShortcutKey = true;
                    if (assign_ShortcutKey)
                    {
                        var path = Path.Combine(FolderPath, typeof(MyPlugin).Namespace + ".cuix");
                        Debug.Print("path=" + path);

                        string sMainCuiFile = (string)Application.GetSystemVariable("MENUNAME");
                        sMainCuiFile += ".cuix";
                        var oCs = new CustomizationSection(sMainCuiFile);

                        var ps = oCs.PartialCuiFiles;

                        if (File.Exists(path))
                        {
                            if (!ps.Contains(path)) Application.LoadPartialMenu(path.Replace("\\", "/"));
                            return;
                        }

                        var newoCs = new CustomizationSection();
                        newoCs.MenuGroupName = typeof(MyPlugin).Namespace;

                        MacroGroup mg = newoCs.MenuGroup.MacroGroups.Count > 0 ? newoCs.MenuGroup.MacroGroups[0] : new MacroGroup("ID_Debuger_MacroGroup", newoCs.MenuGroup);
                        MenuMacro mcr = new MenuMacro(mg, "Copy All F4", "^C^C_F4", "ID_Debuger_CMD_F4");
                        var newover = new MenuAccelerator(mcr, "F4", newoCs.MenuGroup);

                        mcr = new MenuMacro(mg, "ProgressMeter F4", "^C^C_F5", "ID_Debuger_CMD_F5");
                        newover = new MenuAccelerator(mcr, "F5", newoCs.MenuGroup);

                        mcr = new MenuMacro(mg, "COPYHIST F6", "^C^C_F6", "ID_Debuger_CMD_F6");
                        newover = new MenuAccelerator(mcr, "F6", newoCs.MenuGroup);

                        newoCs.SaveAs(path);
                        Debug.Print("Saveas " + path);

                        //ps.Add(path);
                        Application.LoadPartialMenu(path.Replace("\\", "/"));
                    }
                }
            }
            else
            {
                Debug.Print("Tab=null + ID_ADDINSTAB + RBN_00012112", Caption);
                foreach (var t in ribbonControl.Tabs)
                {
                    Debug.Print(" tab Title=" + t.Title + " Id=" + t.Id + " Name=" + t.Name + " AutomationName=" + t.AutomationName);
                }
            }
        }

        public static RibbonButton F4 = null;
        public static RibbonButton F5 = null;
        public static RibbonButton F6 = null;
    }
}