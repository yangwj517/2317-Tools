
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ModbusConfigPlugin.pojo
{
    public class Item
    {
        protected string DataType { get; set; }

        protected string Name { get; set; }

        public  string Address { get; set; }

        protected string Description { get; set; }

        public Item() { }

        public Item(string dataType, string name, string address, string description)
        {
            DataType = dataType;
            Name = name;
            Address = address;
            Description = description;
        }

        public virtual XElement ToXElement()
        {
            var element = new XElement("Item",
                new XAttribute("DataType", DataType),
                new XAttribute("Address", Address),
                new XAttribute("Name", Name),
                new XAttribute("Description", Description)   
            );
            return element;
        }

    }
}
