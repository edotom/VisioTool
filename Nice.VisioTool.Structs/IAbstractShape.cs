using System.Collections.Generic;

namespace Nice.VisioTool.Structs
{
    public interface IAbstractShape
    {
        ICollection<string> ConnectedFrom { get; set; }
        ICollection<string> ConnectedTo { get; set; }
        string Name { get; set; }
        string Type { get; set; }
        int UniqueID { get; set; }
    }
}