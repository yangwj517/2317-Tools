using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModbusConfigPlugin.pojo.Impl
{
    internal class FunCode_5 : Item
    {

        public FunCode_5() { }

        public FunCode_5(string Name, string Address, string Description, string DataType)
        {
            this.Name = Name;
            this.Address = Address;
            this.Description = Description;
            this.DataType = DataType;
        }

    }
}
