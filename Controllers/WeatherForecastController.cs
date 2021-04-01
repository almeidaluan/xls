using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace xls.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        public void Get()
        {
            ExcelPackage excel = new ExcelPackage();
            var testData = new List<object[]>();

            excel.Workbook.Worksheets.Add("Pu Ativo");
            excel.Workbook.Worksheets.Add("PU SWAP");
            excel.Workbook.Worksheets.Add("PU Compromissada");


            List<EnvioSwap> envioList = new List<EnvioSwap>();
            Console.WriteLine("comecou o fluxo");

            EnvioSwap swap = new EnvioSwap();
            swap.CodSwap = 1234;
            swap.DtReference = new DateTime(2021,12,12);
            swap.ValorAtivo = 123456;
            swap.ValorPassivo = 678910;
            swap.TipoCotacao = "F";
            swap.Observacao  = "Importado com Sucesso";

            var excelWorksheet = excel.Workbook.Worksheets["PU SWAP"];

            string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";

            excelWorksheet.Cells[headerRange].LoadFromArrays(headerRow);
            

            for (int i = 0; i < 10; i++)
            {
                envioList.Add(swap);
            }

            Console.WriteLine(JsonConvert.SerializeObject(envioList));
            
            foreach (var t in envioList)
            {
                Console.WriteLine(t.CodSwap);
                testData.Add( new object[]{t.CodSwap, t.DtReference, t.ValorAtivo, t.ValorPassivo, t.TipoCotacao, t.Observacao});
               
            }
           excelWorksheet.Cells[$"B2:B${testData.Count + 1}"].Style.Numberformat.Format = "ddd-mm-dd";
           excelWorksheet.Cells[2, 1].LoadFromArrays(testData);    
            

            FileInfo excelFile = new FileInfo(@"C:\Users\almei\Documents\teste-excel\test.xlsx");
            excel.SaveAs(excelFile);
        }

          public class EnvioSwap{
            public int CodSwap { get; set;}
            public DateTime DtReference { get; set;}
            public decimal ValorAtivo { get; set;}
            public decimal ValorPassivo { get; set;}
            public string TipoCotacao { get; set;}
            public string Observacao { get; set;}
         }

        List<string[]> headerRow = new List<string[]>()
        {
            new string[] { "Cod Swap", "Data", "Valor Ativo", "Valor Passivo", "Tipo Cotacao", "Observacao" }
        };

         List<EnvioSwap[]> teste = new List<EnvioSwap[]>()
        {
            
        };

    }
  
}
