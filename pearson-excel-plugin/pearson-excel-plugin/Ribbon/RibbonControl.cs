using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace Pearson.Excel.Plugin.Ribbon
{
    public class RibbonControl
    {
        private readonly IRibbonUI _ribbon;
        private readonly string _type;
        public string Id { get; set; }
        public string Label { get; set; }
        public string ImageMso { get; set; }
        public bool IsEnabled { get; set; }
        public bool IsVisible { get; set; }
        public Action OnInvalidate { get; set; }

        public RibbonControl(string type, IRibbonUI ribbon)
        {
            _type = type;
            _ribbon = ribbon;
        }

        public RibbonControl(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
        }

        public void Invalidate()
        {
            OnInvalidate?.Invoke();
            ExcelAsyncUtil.QueueAsMacro(() => _ribbon.InvalidateControl(Id));
        }

        #region For dynamic stuff

        protected virtual List<Tuple<string, string>> GetAttributes()
        {
            var attributes = new List<Tuple<string, string>>
            {
                new Tuple<string, string>("id", Id),
                new Tuple<string, string>("label", Label),
            };
            return attributes;
        }

        public XElement GetXml()
        {
            var attributes = GetAttributes();
            XNamespace ns = ExcelRibbon.NamespaceCustomUI2010;
            var element = new XElement(ns + _type);
            attributes.ForEach(att => element.Add(new XAttribute(att.Item1, att.Item2)));
            return element;
        }

        #endregion
    }
}