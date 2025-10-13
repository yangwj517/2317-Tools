using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ModbusConfigPlugin.pojo.Impl
{
    internal class FunCode_1 : Item
    {
        string ScadaScanner {  get; set; }

        public FunCode_1() { }

        public FunCode_1(string Name, string Address, string Description, string DataType, string scadaScanner)
        {
            this.Name = Name;
            this.Address = Address;
            this.Description = Description;
            this.DataType = DataType;
            this.ScadaScanner = scadaScanner;
        }


        public override XElement ToXElement()
        {
            var element = new XElement("Item",
                new XAttribute("DataType", DataType),
                new XAttribute("Address", Address),
                new XAttribute("ScadaScanner", ScadaScanner),
                new XAttribute("Name", Name),
                new XAttribute("Description", Description)
            );
            return element;
        }

    }
}
