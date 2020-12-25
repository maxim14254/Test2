using System.Runtime.Serialization;


namespace ConsoleApp1
{
    [DataContract]
    class Pump
    {
        [DataMember]
        public string Code { get; set; }
        [DataMember]
        public string Type { get; set; }
        [DataMember]
        public string SubType { get; set; }
        [DataMember]
        public int SafetyClass { get; set; }
        [DataMember]
        public int D { get; set; }
        [DataMember]
        public string Fluid { get; set; }
        [DataMember]
        public double Weight { get; set; }
    }
}
