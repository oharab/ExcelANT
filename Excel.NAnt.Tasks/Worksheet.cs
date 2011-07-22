namespace Excel.NAnt.Tasks
{
    using System;
    using global::NAnt.Core;
    using global::NAnt.Core.Attributes;
    using global::NAnt.Core.Types;

    [ElementName("worksheet")]
    public class Worksheet:Element
    {


        [TaskAttribute("name")]
        [StringValidator(AllowEmpty = false)]  
        public string SheetName { get; set; }
    }
}
