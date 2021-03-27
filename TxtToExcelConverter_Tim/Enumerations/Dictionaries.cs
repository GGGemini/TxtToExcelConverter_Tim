using System.Collections.Generic;

namespace TxtToExcelConverter_Tim.Enumerations
{
    public static class Dictionaries
    {
        public static Dictionary<string, string> ComponentTypes
        { 
            get
            {
                return new Dictionary<string, string>
                {
                    { "BQ", "(Quartz Resonator) 石英谐振器" },
                    { "C", "(Сondenser) 电容" },
                    { "R", "(Resistor) 电阻" },
                    { "VT", "(Transistor) 晶体管" },
                    { "D", "(Chip) 芯片" },
                    { "VD", "(Diode) 二极管" },
                    { "HL", "(LED) 发光二极管" },
                    { "XS", "(Connector) 连接器" },
                    { "XT", "(Connector) 连接器" },
                    { "XP", "(Connector) 连接器" },
                    { "J", "(Jumper wire)" }
                };
            }
        }
    }
}