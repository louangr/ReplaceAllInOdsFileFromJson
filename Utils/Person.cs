namespace Utils
{
    using System.Runtime.Serialization;
    
    [DataContract]
    public class Person
    {

        [DataMember(Name = "name")]
        public string Name { get; set; }

        [DataMember(Name = "id")]
        public int Id { get; set; }
    }
}
