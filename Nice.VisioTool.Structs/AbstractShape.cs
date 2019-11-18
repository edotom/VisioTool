using System;
using System.Collections.Generic;

namespace Nice.VisioTool.Structs
{
    public class AbstractShape : IAbstractShape
    {
        private ICollection<string> _connectedTo;

        public int UniqueID { get; set; }
        public string Name { get; set; }
        public ICollection<string> ConnectedFrom { get; set; }
        public ICollection<string> ConnectedTo { get => _connectedTo; set => _connectedTo = value; }
        public string Type { get; set; }

        public AbstractShape()
        {
            UniqueID = 0;
            Name = String.Empty;
            Type = String.Empty;
            ConnectedFrom = new List<string>();
            _connectedTo = new List<string>();
        }
    }
}




