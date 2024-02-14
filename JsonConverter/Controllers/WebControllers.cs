using Microsoft.AspNetCore.Mvc;
using JsonConverter.model;
using OfficeOpenXml;

namespace JsonConverter.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class WebControllers : Controller
    {
        [HttpPost]
        public ActionResult<List<User>> UploadFile([FromBody] User newUser)
        {
            var list = new List<User>();
            //Chamba del codigo del controller que hace lo del Excel.

            //var result = await sonConvert.DeserializeObject<WeatherForecast>(json);

            //TEST
            var word = newUser.CorreoElectronico + " TEST " + newUser.NombreEgresado;

            var response = new User();

            response.NombreEgresado = word;

            var content = ReadExcelFile(@"C:\Users\re3ne\OneDrive\Escritorio\Excel\Sistemas.xlsx", out List<User> listUsers);

            list.Add(response);

            return Ok(listUsers);
            
        }

        private bool ReadExcelFile(string filePath, out List<User> egresados)
        {
            bool result = false;
            egresados = new List<User>();

            if (!string.IsNullOrEmpty(filePath))
            {
                string fileExtension = Path.GetExtension(filePath);

                if (fileExtension == ".xls" || fileExtension == ".xlsx")
                {
                    var fileLocation = new FileInfo(filePath);

                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    using (ExcelPackage package = new ExcelPackage(fileLocation))
                    {
                        if (package.Workbook.Worksheets.Any())
                        {
                            ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
                            int totalRows = workSheet.Dimension.Rows;

                            for (int i = 2; i <= totalRows; i++)
                            {
                                //NumeroControl   NombreEgresado  Telefono    CorreoElectronico     Generacion
                                egresados.Add(new User
                                {
                                    EgresadoId = workSheet.Cells[i, 1].Value.ToString(),
                                    NombreEgresado = workSheet.Cells[i, 2].Value.ToString(),
                                    Telefono = workSheet.Cells[i, 3].Value.ToString(),
                                    CorreoElectronico = workSheet.Cells[i, 4].Value.ToString(),
                                    Generacion = workSheet.Cells[i, 5].Value.ToString(),
                                    CarreraId = 5
                                });
                            }
                            result = true;
                        }
                    }
                }
            }
            return result;
        }



        //public bool InsertaInfoFromExcel(string path, int carreraId, out List<UsuarioEgresado> egresados)
        //{
        //    bool result = false;
        //    _datos.ReadExcelFile(path, carreraId, out egresados);
        //    //implementar guardado a la bd
        //    foreach (UsuarioEgresado egresado in egresados)
        //    {
        //        var resultado = _datos.InsertarEgresado(egresado);
        //        if (resultado)
        //        {
        //            Console.WriteLine("insertado con exito");
        //            result = true;
        //        }
        //        else
        //        {
        //            Console.WriteLine("No se pudo insertar " + egresado.NombreEgresado);
        //        }
        //    }

        //    return result;
        //}
    }
}
