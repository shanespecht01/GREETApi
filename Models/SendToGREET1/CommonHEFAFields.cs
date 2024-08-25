namespace GREETApi.Models.SendToGREET1
{
    public class CommonHEFAFields : CommonFields
    {
        public decimal HEFA_feed_usage { get; set; }
        public decimal T_D_HEFA_SAF_Barge_Distance { get; set; }
        public decimal T_D_HEFA_SAF_Barge_Share { get; set; }
        public decimal T_D_HEFA_SAF_Pipeline_Distance { get; set; }
        public decimal T_D_HEFA_SAF_Pipeline_Share { get; set; }
        public decimal T_D_HEFA_SAF_Rail_Distance { get; set; }
        public decimal T_D_HEFA_SAF_Rail_Share { get; set; }
        public decimal T_D_HEFA_SAF_Truck_Distance { get; set; }
    }
}
