using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ModbusConfigPlugin.pojo.Impl
{
    internal class FunCode_6 : Item
    {

        string SwapBytes { get; set; }

        string SwapWords { get; set; }

        string Deadband { get; set; }

        string OutputConversion { get; set; }

        public FunCode_6() { }

        public FunCode_6(string Name, string Address, string Description, string DataType,string swapBytes, string swapWords ,string deadband,string outputConversion)
        {
            this.Name = Name;
            this.Address = Address;
            this.Description = Description;
            this.DataType = DataType;
            SwapBytes = swapBytes;
            SwapWords = swapWords;
            Deadband = deadband;
            OutputConversion = outputConversion;
            
        }

        public override XElement ToXElement()
        {
            var element = new XElement("Item",
                new XAttribute("DataType", DataType),
                new XAttribute("Address", Address),
                new XAttribute("SwapBytes", SwapBytes),
                new XAttribute("SwapWords", SwapWords),
                new XAttribute("OutputConversion", OutputConversion),
                new XAttribute("Name", Name),
                new XAttribute("Description", Description),
                new XAttribute("Deadband", Deadband)
            );
            return element;
        }
    }
}
