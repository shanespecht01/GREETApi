using GREETApi.Services;
using GREETApi.Models.SendToGREET1;
using Microsoft.AspNetCore.Mvc;
using DocumentFormat.OpenXml.Office2016.Excel;

namespace GREETApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class GREETController : ControllerBase
    {
        private readonly IExcelService _excelService;

        public GREETController()
        {
            _excelService = new ExcelService();
        }

        public enum DataType
        {
            CanolaHEFA,
            CornATJethanol,
            DistillersCornOilHEFA,
            SoybeanHEFA,
            SugarcaneATJethanol,
            TallowHEFA,
            UCOHEFA
        }

        [HttpPost("send-CanolaHEFA-data")]
        public IActionResult SendCanolaHEFAData(CanolaHEFA data)
        {
            try
            {
                var filePath = "Excel/GREET1_2023_Rev1.xlsm";
                var sheetName = "GREET Connection";

                _excelService.SendToGREET1(filePath, sheetName, data);
                return Ok("Data updated successfully.");
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("send-CornATJethanol-data")]
        public IActionResult SendCornATJethanolData(CornATJethanol data)
        {
            try
            {
                var filePath = "Excel/GREET1_2023_Rev1.xlsm";
                var sheetName = "GREET Connection";

                _excelService.SendToGREET1(filePath, sheetName, data);
                return Ok("Data updated successfully.");
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("send-DistillersCornOilHEFA-data")]
        public IActionResult SendDistillersCornOilHEFAData(DistillersCornOilHEFA data)
        {
            try
            {
                var filePath = "Excel/GREET1_2023_Rev1.xlsm";
                var sheetName = "GREET Connection";

                _excelService.SendToGREET1(filePath, sheetName, data);
                return Ok("Data updated successfully.");
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("send-SoybeanHEFA-data")]
        public IActionResult SendSoybeanHEFAData(SoybeanHEFA data)
        {
            try
            {
                var filePath = "Excel/GREET1_2023_Rev1.xlsm";
                var sheetName = "GREET Connection";

                _excelService.SendToGREET1(filePath, sheetName, data);
                return Ok("Data updated successfully.");
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("send-SugarcaneATJethanol-data")]
        public IActionResult SendSugarcaneATJethanolData(SugarcaneATJethanol data)
        {
            try
            {
                var filePath = "Excel/GREET1_2023_Rev1.xlsm";
                var sheetName = "GREET Connection";

                _excelService.SendToGREET1(filePath, sheetName, data);
                return Ok("Data updated successfully.");
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("send-TallowHEFA-data")]
        public IActionResult SendTallowHEFAData(TallowHEFA data)
        {
            try
            {
                var filePath = "Excel/GREET1_2023_Rev1.xlsm";
                var sheetName = "GREET Connection";

                _excelService.SendToGREET1(filePath, sheetName, data);
                return Ok("Data updated successfully.");
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("send-UCOHEFA-data")]
        public IActionResult SendUCOHEFAData(UCOHEFA data)
        {
            try
            {
                var filePath = "Excel/GREET1_2023_Rev1.xlsm";
                var sheetName = "GREET Connection";

                _excelService.SendToGREET1(filePath, sheetName, data);
                return Ok("Data updated successfully.");
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("send-data")]
        public IActionResult SendData([FromQuery] DataType dataType, [FromBody] object data)
        {
            try
            {
                var filePath = "Excel/GREET1_2023_Rev1.xlsm";
                var sheetName = "GREET Connection";

                switch (dataType)
                {
                    case DataType.CanolaHEFA:
                        if (data is not CanolaHEFA canolaHEFAData)
                            return BadRequest(new BadHttpRequestException("Invalid data for CanolaHEFA"));
                        break;

                    case DataType.CornATJethanol:
                        if (data is not CornATJethanol cornATJethanolData)
                            return BadRequest(new BadHttpRequestException("Invalid data for CornATJethanol"));
                        break;

                    case DataType.DistillersCornOilHEFA:
                        if (data is not DistillersCornOilHEFA distillersCornOilHEFAData)
                            return BadRequest(new BadHttpRequestException("Invalid data for DistillersCornOilHEFA"));
                        break;

                    case DataType.SoybeanHEFA:
                        if (data is not SoybeanHEFA soybeanHEFAData)
                            return BadRequest(new BadHttpRequestException("Invalid data for SoybeanHEFA"));
                        break;

                    case DataType.SugarcaneATJethanol:
                        if (data is not SugarcaneATJethanol sugarcaneATJethanolData)
                            return BadRequest(new BadHttpRequestException("Invalid data for SugarcaneATJethanol"));
                        break;

                    case DataType.TallowHEFA:
                        if (data is not TallowHEFA tallowHEFAData)
                            return BadRequest(new BadHttpRequestException("Invalid data for TallowHEFA"));
                        break;

                    case DataType.UCOHEFA:
                        if (data is not UCOHEFA ucoHEFAData)
                            return BadRequest(new BadHttpRequestException("Invalid data for UCOHEFA"));
                        break;

                    default:
                        return BadRequest("Invalid data type.");
                }

                _excelService.SendToGREET1(filePath, sheetName, data);

                return Ok("Data updated successfully.");
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpGet("get-data")]
        public IActionResult GetData()
        {
            try
            {
                var filePath = "Excel/GREET1_2023_Rev1.xlsm";
                var sheetName = "JetFuel_WTP";
                var data = _excelService.GetFromGREET1(filePath, sheetName);
                return Ok(data);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
    }
}
