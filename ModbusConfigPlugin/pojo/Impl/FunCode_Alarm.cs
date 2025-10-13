using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ModbusConfigPlugin.pojo.Impl
{
    internal class FunCode_Alarm : Item
    {
        string ScadaScanner {  get; set; }

        string AckStatus { get; set; }

        string AckCommand { get; set; }

        public FunCode_Alarm() {

        }
        public FunCode_Alarm(string DataType , string Name ,String Adderss ,String Description , string scadaScanner, string ackStatus, string ackCommand)
        {
            this.DataType = DataType;
            this.Name = Name;
            this.Address= Adderss;
            this.Description = Description;
            ScadaScanner = scadaScanner;
            AckStatus = ackStatus;
            AckCommand = ackCommand;
        }

        public override XElement ToXElement()
        {
            var element = new XElement("Item",
                new XAttribute("DataType", DataType),
                new XAttribute("Address", Address),
                new XAttribute("ScadaScanner", ScadaScanner),
                new XAttribute("Name", Name),
                new XAttribute("Description", Description),
                new XAttribute("AckStatus", AckStatus),
                new XAttribute("AckCommand", AckCommand)
            );
            return element;
        }
    }
}
