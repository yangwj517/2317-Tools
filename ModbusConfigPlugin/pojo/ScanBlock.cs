using DocumentFormat.OpenXml.Office.CustomUI;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ModbusConfigPlugin.pojo
{
    internal class ScanBlock
    {

        // 扫描块点名
        private string Name {  get; set; }

        // NamePartner
        private string NamePartner;

        // Description
        private string Description;

        // DescriptionPartner
        private string DescriptionPartner;

        private string Type;  // 默认给定

        // 根据功能码不同给定
        private string Operation;

        //以下均为默认给定参数
        private string ScanInterval;
        private string IntervalUnits;
        private string FastScanInterval;
        private string FastScanCount;
        private string ScanOffset;
        private string DemandScanOffset;
        private string TriggerPoint;
        private string TriggerPointPartner;
        private string ScanInhibitPoint ;
        private string ScanInhibitPointPartner;
        private string InhibitInitialExceptionOutput;
        private string InhibitInitialExceptionTime;

        // 扫描块IO数据存储列表
        public List<Item>  Items { get; set; }

        // 每个引用块存储的最大数量
        private int MaxNum { get; set; }


        public ScanBlock(int index ,string Operation, int maxNum)
        {
            this.Operation = Operation;
            Name = "$ScanBlock"+index;
            MaxNum = maxNum;
            NamePartner = "";
            Description = "";
            DescriptionPartner = "";
            Type = "Periodic";
            ScanInterval = "5";
            IntervalUnits = "Seconds";
            FastScanInterval = "1";
            FastScanCount = "0";
            ScanOffset = "0";
            DemandScanOffset = "0";
            TriggerPoint = "";
            TriggerPointPartner = "";
            ScanInhibitPoint = "";
            ScanInhibitPointPartner = "";
            InhibitInitialExceptionOutput = "";
            InhibitInitialExceptionTime = "0";
        }

        public XElement ToXElement()
        {
            var element = new XElement("ScanBlock",
                new XAttribute("Name", Name),
                new XAttribute("NamePartner", NamePartner),
                new XAttribute("Description", Description),
                new XAttribute("DescriptionPartner", DescriptionPartner),
                new XAttribute("Type", Type),
                new XAttribute("Operation", Operation),
                new XAttribute("ScanInterval", ScanInterval),
                new XAttribute("IntervalUnits", IntervalUnits),
                new XAttribute("FastScanInterval", FastScanInterval),
                new XAttribute("FastScanCount", FastScanCount),
                new XAttribute("ScanOffset", ScanOffset),
                new XAttribute("DemandScanOffset", DemandScanOffset),
                new XAttribute("TriggerPoint", TriggerPoint),
                new XAttribute("TriggerPointPartner", TriggerPointPartner),
                new XAttribute("ScanInhibitPoint", ScanInhibitPoint),
                new XAttribute("ScanInhibitPointPartner", ScanInhibitPointPartner),
                new XAttribute("InhibitInitialExceptionOutput", InhibitInitialExceptionOutput),
                new XAttribute("InhibitInitialExceptionTime", InhibitInitialExceptionTime)
            );

            foreach (var item in Items)
            {
                element.Add(item.ToXElement());
            }
            return element;
        }
    }
}
