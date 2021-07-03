using System.Web.Http;

namespace TaskExcelApi.Controllers
{
    public class ExcelController : ApiController
    {
        [HttpPost]
        public Response GetExcel(ExcelModel modelData)
        {
            Response response = cDAL.ReadExcel(modelData);
            return response;
        }
    }
}