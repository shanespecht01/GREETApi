namespace GREETApi.Models.SendToGREET1
{
    public class CommonETJFields : CommonFields
    {
        public decimal ETJ_Standalone_Electricity { get; set; }
        public decimal ETJ_Standalone_Ethanol { get; set; }
        public decimal ETJ_Standalone_Hydrogen { get; set; }
        public decimal ETJ_Standalone_NG { get; set; }
        public decimal Ethanol_ETJ_distributed_Catalysts { get; set; }
        public decimal Green_Ammonia_Share { get; set; }
        public decimal T_D_ETJ_SAF_Barge_Distance { get; set; }
        public decimal T_D_ETJ_SAF_Barge_Share { get; set; }
        public decimal T_D_ETJ_SAF_Pipeline_Distance { get; set; }
        public decimal T_D_ETJ_SAF_Pipeline_Share { get; set; }
        public decimal T_D_ETJ_SAF_Rail_Distance { get; set; }
        public decimal T_D_ETJ_SAF_Rail_Share { get; set; }
        public decimal T_D_ETJ_SAF_Truck_Distance { get; set; }
        public decimal T_D_ETJ_SAF_Truckl_Share { get; set; }
    }
}
