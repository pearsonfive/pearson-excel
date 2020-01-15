using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using NetOffice.ExcelApi;
using Pearson.Excel.Plugin.Ribbon;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml.Linq;
using Pearson.Excel.Plugin.Extensions;
using Action = System.Action;
using Application = NetOffice.ExcelApi.Application;

namespace Pearson.Excel.Plugin
{
    [ComVisible(true)]
    public class RibbonHandler : ExcelRibbon
    {
        private static readonly Dictionary<string, RibbonControl> controls = new Dictionary<string, RibbonControl>();

        private static readonly Application app = new Application(null, ExcelDnaUtil.Application);
        private IRibbonUI _ribbon;

        private XElement _dynamicMenuContent;

        public void Ribbon_Load(IRibbonUI sender)
        {
            _ribbon = sender;

            new List<RibbonControl>
            {
                new RibbonControl(_ribbon)
                {
                    Id = "CalculationGroup",
                    IsVisible = true
                },
                new ButtonRibbonControl(_ribbon)
                {
                    Id = "btnCalcNow",
                    Label = "Calculate Now",
                    ImageMso = "AcceptInvitation",
                    IsEnabled = true,
                    IsVisible = true,
                    OnAction = control => app.Calculate()
                },
                new ButtonRibbonControl(_ribbon)
                {
                    Id = "btnCalcSheet",
                    Label = "Calculate Sheet",
                    ImageMso = "AccessFormDatasheet",
                    IsEnabled = true,
                    IsVisible = true,
                    OnAction = control =>
                    {
                        ((Worksheet) app.ActiveSheet)?.Calculate();
                    }
                },
                new RibbonControl(_ribbon)
                {
                    Id = "ExamplesGroup",
                    IsVisible = true
                },
                new DropDownRibbonControl(_ribbon)
                {
                    Id="dropDownExample",
                    IsEnabled = true,
                    SelectedItemId = "i2",
                    OnAction = control => MessageBox.Show($"You selected {control.SelectedItemId}")
                },
                new RibbonControl(_ribbon)
                {
                    Id="dynamicMenuExample",
                    IsEnabled = true,
                    Label="Dynamic",
                    ImageMso = "ServerConnection"
                }

            }.ForEach(control => controls[control.Id] = control);

            // dynamic menu content
            XNamespace ns = NamespaceCustomUI2010;
            _dynamicMenuContent = new XElement(ns + "menu");
            buildDynamicContent(_dynamicMenuContent, ns);
        }

        private void buildDynamicContent(XElement menu, XNamespace ns)
        {
            var counter = 0;

            new[] { "Folder_1", "Folder_2", "Folder_3" }.ForEach(folder =>
              {
                  counter++;
                  var menuItem = buildMenuItem(ns, folder, $"f{counter}");
                  menu.Add(menuItem);

                  new[] { "Item_1", "Item_2", "Item_3" }.ForEach(item =>
                  {
                      counter++;
                      createButton(
                          () => MessageBox.Show($@"{folder}\{item}"),
                          $"dynamicButton{counter}",
                          item,
                          "",
                          menuItem);
                  });
              });
        }

        private void createButton(Action action, string id, string label, string imageMso, XElement menu)
        {
            var b = new ButtonRibbonControl(_ribbon)
            {
                Id = id,
                Label = label,
                ImageMso = imageMso,
                OnAction = control =>
                {
                    action();
                }
            };
            controls[id] = b;
            menu.Add(b.GetXml());
        }

        private XElement buildMenuItem(XNamespace ns, string label, string id)
        {
            return new XElement(ns + "menu", new XAttribute("label", label), new XAttribute("id", id));
        }

        private void invalidateRibbon()
        {
            ExcelAsyncUtil.QueueAsMacro(() => _ribbon.Invalidate());
        }

        #region Get*****

        public string GetLabel(IRibbonControl control)
        {
            var c = controls[control.Id];
            return c.Label;
        }

        public string GetImage(IRibbonControl control)
        {
            var c = controls[control.Id];
            return c.ImageMso;
        }

        public bool GetVisible(IRibbonControl control)
        {
            var c = controls[control.Id];
            return c.IsVisible;
        }

        public bool GetEnabled(IRibbonControl control)
        {
            var c = controls[control.Id];
            return c.IsEnabled;
        }

        public string GetSelectedItemId(IRibbonControl control)
        {
            var c = (DropDownRibbonControl)controls[control.Id];
            return c.SelectedItemId;
        }

        public string GetDynamicMenuContent(IRibbonControl control)
        {
            return _dynamicMenuContent.ToString();
        }

        #endregion

        #region On***

        public void OnAction(IRibbonControl control)
        {
            var c = (ButtonRibbonControl)controls[control.Id];
            c.OnAction(control);
        }

        public void OnActionDropDown(IRibbonControl control, string selectedId, int selectedIndex)
        {
            var c = (DropDownRibbonControl)controls[control.Id];
            c.SelectedItemId = selectedId;
            c.SelectedIndex = selectedIndex;
            c.OnAction(c);
        }

        #endregion
    }
}