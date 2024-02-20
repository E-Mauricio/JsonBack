using Microsoft.AspNetCore.Mvc;
using JsonConverter.model;
using OfficeOpenXml;
using System.Linq;
using Microsoft.Extensions.Options;

namespace JsonConverter.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class WebControllers : Controller
    {
        private readonly IWebHostEnvironment _host;
        public WebControllers(IWebHostEnvironment host)
        {
            _host = host;
        }

        [HttpPost]
        public ActionResult<List<DataTest>> UploadFile()
        {

            //SAve file to wwwroot
            var content = TestReadExcelFile(@"C:\Users\re3ne\OneDrive\Escritorio\Excel\archivo muestra_Json.xlsx", out List<DataTest> jsonRecords);

            return Ok(jsonRecords);

        }

        [HttpPost("Ping")]
        public ActionResult<string> Pong([FromBody] string name)
        {
            return "Pong" + name;
        }

        [HttpPost("SubirArchivo")]
        [DisableRequestSizeLimit]
        public async Task<ActionResult<List<DataTest>>> SubirArchivo([FromForm] IFormFile file)
        {

            List<DataTest> json = new List<DataTest>();

            // full path to file in temp location
            try
            {

                var filePath = Path.Combine(_host.WebRootPath, "Excel", file.FileName);
                if (file.Length > 0)
                {
                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        await file.CopyToAsync(stream);
                    }
                }

                if (TestReadExcelFile(filePath, out json))
                {
                    return Ok(json);
                }
                else
                {
                    return BadRequest("Formato incompatible.");
                }

            }
            catch (Exception ex)
            {
                return BadRequest(new { ex.Message, ex.StackTrace });
            }
        }

        private bool TestReadExcelFile(string filePath, out List<DataTest> entries)
        {
            bool result = false;
            entries = new List<DataTest>();

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

                            int totalRows = workSheet.Dimension.Rows - 1;

                            //Quitar lo hardcodeado
                            int initialRow = 0;
                            int column = 1;

                            bool contentType = false;

                            for (int i = 7; i <= totalRows; i++)
                            {
                                entries.Add(new DataTest
                                {

                                    //Date = workSheet.Cells[i, 1].Text,
                                    //Date = workSheet.Cells[i, 1].Value.ToString() ?? " ",
                                    Date = (workSheet.Cells[i, 1].Value ?? " ").ToString(),

                                    InitialInventory = (workSheet.Cells[i, 2].Value ?? " ").ToString(),

                                    Entradas = new Entradas()
                                    {
                                        Compras = (workSheet.Cells[i, 3].Value ?? " ").ToString(),
                                        Traspaso = (workSheet.Cells[i, 4].Value ?? " ").ToString(),
                                        Recargas = (workSheet.Cells[i, 5].Value ?? " ").ToString(),
                                        Ajuste = (workSheet.Cells[i, 6].Value ?? " ").ToString(),
                                        TotalEntradas = (workSheet.Cells[i, 7].Value ?? " ").ToString(),
                                    },

                                    Salidas = new Salidas()
                                    {
                                        Publico = (workSheet.Cells[i, 8].Value ?? " ").ToString(),
                                        Traspaso = (workSheet.Cells[i, 9].Value ?? " ").ToString(),
                                        Recargas = (workSheet.Cells[i, 10].Value ?? " ").ToString(),
                                        Ajuste = (workSheet.Cells[i, 11].Value ?? " ").ToString(),
                                        Liquidacion = (workSheet.Cells[i, 12].Value ?? " ").ToString(),
                                        TotalSalidas = (workSheet.Cells[i, 13].Value ?? " ").ToString(),
                                    },

                                    FinalInventory = (workSheet.Cells[i, 14].Value ?? " ").ToString(),
                                    FisicosCapturados = (workSheet.Cells[i, 15].Value ?? " ").ToString(),

                                    Inventario = new Inventario()
                                    {
                                        Dia = (workSheet.Cells[i, 16].Value ?? " ").ToString(),
                                        Acumulado = (workSheet.Cells[i, 17].Value ?? " ").ToString()
                                    }

                                });
                            }

                            result = true;
                        }
                    }
                }
            }
            return result;
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
