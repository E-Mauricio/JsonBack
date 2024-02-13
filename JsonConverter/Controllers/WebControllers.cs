using Microsoft.AspNetCore.Mvc;

namespace JsonConverter.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class WebControllers : Controller
    {
        [HttpPost]
        public IActionResult UploadFile(string nameTest)
        {
            //Chamba del codigo del controller que hace lo del Excel.

            //var result = await sonConvert.DeserializeObject<WeatherForecast>(json);

            //TEST
            var word = nameTest + " TEST " + nameTest;

            return Ok(word);
            
        }
    }
}
