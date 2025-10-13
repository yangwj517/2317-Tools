using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ModbusConfigPlugin.pojo.Impl
{
    internal class FunCode_4 : Item
    {
        string ScadaScanner { get; set; }

        string SwapBytes { get; set; }

        string SwapWords { get; set; }

        public FunCode_4() { }

        public FunCode_4(string Name, string Address, string Description, string DataType, string scadaScanner, string swapBytes, string swapWords)
        {
            this.Name = Name;
            this.Address = Address;
            this.Description = Description;
            this.DataType = DataType;
            ScadaScanner = scadaScanner;
            SwapBytes = swapBytes;
            SwapWords = swapWords;
        }

        public override XElement ToXElement()
        {
            var element = new XElement("Item",
                new XAttribute("DataType", DataType),
                new XAttribute("Address", Address),
                new XAttribute("SwapBytes", SwapBytes),
                new XAttribute("SwapWords", SwapWords),
                new XAttribute("ScadaScanner", ScadaScanner),
                new XAttribute("Name", Name),
                new XAttribute("Description", Description)
            );
            return element;
        }
    }
}
