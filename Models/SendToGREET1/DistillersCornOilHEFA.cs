namespace GREETApi.Models.SendToGREET1
{
    public class DistillersCornOilHEFA : CommonHEFAFields
    {
        public decimal CornOil_HEFA_Electricity { get; set; }
        public decimal CornOil_HEFA_Hydrogen { get; set; }
        public decimal CornOil_HEFA_Natural_gas { get; set; }
        public decimal HEFA_Allocation { get; set; }
    }
}
