using System;

namespace MondayMaster
{
    public class UpdateRecord
    {
        public string Name { get; set; }
        public DateTime? ExitDateCurrent { get; set; }
        public DateTime? ExitDateOriginal { get; set; }
        public string Comment { get; set; }
        public string Health { get; set; }
        public string Id { get; set; }
        public string Header { get; set; }
        public string ProductManager { get; internal set; }
        public string LeadEng { get; internal set; }

        public override string ToString()
        {
            return $"[{Header}] {Health}: {Name}";
        }
    }
}