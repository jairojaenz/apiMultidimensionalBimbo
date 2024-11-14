using Microsoft.AspNetCore.Mvc;
using Microsoft.AnalysisServices.AdomdClient;
using System.Collections.Generic;

namespace YourNamespace.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class OlapController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public OlapController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        /// <summary>
        /// ///////////////////////////////////////////////////consulta 1////////////////////////////////////////////////////
        /// </summary>
        /// <returns></returns>
        [HttpGet("Venta total de X Producto por mes")]
        public IActionResult GetSalesData()
        {
            string connectionString = _configuration.GetConnectionString("OlapConnection");
            using (AdomdConnection connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT NON EMPTY { [Measures].[Venta Total] } ON COLUMNS, NON EMPTY {  ([Dimension Productos].[Nombre Producto].[Nombre Producto].ALLMEMBERS * [Dim Date].[Month Name].[Month Name].ALLMEMBERS  ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM [Dimenciones Pan Plus] CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                using (AdomdCommand command = new AdomdCommand(query, connection))
                {
                    var result = command.ExecuteCellSet();
                    var jsonResult = TransformToJSON(result);
                    return Ok(jsonResult);
                }
            }
        }

        /// <summary>
        /// ///////////////////////////////////////////////////consulta 2////////////////////////////////////////////////////
        /// </summary>
        /// <returns></returns>

        [HttpGet("Clientes con mayor cantidad de compras")]
        public IActionResult GetTop10ClientsBySales()
        {
            string connectionString = _configuration.GetConnectionString("OlapConnection");
            using (AdomdConnection connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT NON EMPTY { [Measures].[Cantidad Venta] } ON COLUMNS, NON EMPTY { TOPCOUNT( ([Dimension Cliente].[Nombre Cliente].[Nombre Cliente].ALLMEMBERS), 20, [Measures].[Cantidad Venta] ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM [Dimenciones Pan Plus] CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                using (AdomdCommand command = new AdomdCommand(query, connection))
                {
                    var result = command.ExecuteCellSet();
                    var jsonResult = TransformToJSON(result);
                    return Ok(jsonResult);
                }
            }
        }

        /// <summary>
        /// ///////////////////////////////////////////////////consulta 3////////////////////////////////////////////////////
        /// </summary>
        /// <returns></returns>

        [HttpGet("Ventas total de productos por categorías")]
        public IActionResult GetTop7ProductsByCategory()
        {
            string connectionString = _configuration.GetConnectionString("OlapConnection");
            using (AdomdConnection connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT NON EMPTY { [Measures].[Venta Total] } ON COLUMNS, NON EMPTY { TOPCOUNT ([Dimension Productos].[Nombre Producto].[Nombre Producto].ALLMEMBERS * [Dimension Productos].[Categoria Producto].[Categoria Producto].ALLMEMBERS,7 ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM [Dimenciones Pan Plus] CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                using (AdomdCommand command = new AdomdCommand(query, connection))
                {
                    var result = command.ExecuteCellSet();
                    var jsonResult = TransformToJSON(result);
                    return Ok(jsonResult);
                }
            }
        }


        /// <summary>
        /// //////////////////////////////////////////////////////////Consulta 4///////////////////////////////////////////////////////////
        /// </summary>
        /// <returns></returns>                            Revisar el promedio es igual para todas y no lo veo correcto

        [HttpGet("Promedio de ventas diarias por mes")]
        public IActionResult GetPromedio_ventas_diarias()
        {
            string connectionString = _configuration.GetConnectionString("OlapConnection");
            using (AdomdConnection connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                string query = "WITH MEMBER [Measures].[Promedio ventas diarias] AS AVG(EXISTING ( [Dim Date].[Weekday Name].[Weekday Name].Members), [Measures].[Venta Total] )SELECT NON EMPTY { [Measures].[Promedio ventas diarias] } ON COLUMNS, NON EMPTY {  ([Dim Date].[Month name].[Month name].ALLMEMBERS )} DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM [Dimenciones Pan Plus]";

                using (AdomdCommand command = new AdomdCommand(query, connection))
                {
                    var result = command.ExecuteCellSet();
                    var jsonResult = TransformToJSON(result);
                    return Ok(jsonResult);
                }
            }
        }

        /// <summary>
        /// ////////////////////////////////////////////////////Consulta 5 //////////////////////////////////////////////
        /// </summary>
        /// <returns></returns>

        [HttpGet("Ingreso total por mes")]
        public IActionResult GetIngreso_total_mesual()
        {
            string connectionString = _configuration.GetConnectionString("OlapConnection");
            using (AdomdConnection connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT NON EMPTY { [Measures].[Venta Total] } ON COLUMNS, NON EMPTY { ([Dim Date].[Month Name].[Month Name].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM [Dimenciones Pan Plus] CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                using (AdomdCommand command = new AdomdCommand(query, connection))
                {
                    var result = command.ExecuteCellSet();
                    var jsonResult = TransformToJSON(result);
                    return Ok(jsonResult);
                }
            }
        }


        /// <summary>
        /// ////////////////////////////////////////////////////////////consluta 6//////////////////////////////////////////////////////////////
        /// </summary>
        /// <returns></returns>


        [HttpGet("Ventas semanales por año")]
        public IActionResult GetVentas_semaneles_por_ano()
        {
            string connectionString = _configuration.GetConnectionString("OlapConnection");
            using (AdomdConnection connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT NON EMPTY { [Measures].[Venta Total] } ON COLUMNS, NON EMPTY { ([Dim Date].[Year].[Year].ALLMEMBERS * [Dim Date].[Month Name].[Month Name].ALLMEMBERS * [Dim Date].[Week Of Month].[Week Of Month].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM [Dimenciones Pan Plus] CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                using (AdomdCommand command = new AdomdCommand(query, connection))
                {
                    var result = command.ExecuteCellSet();
                    var jsonResult = TransformToJSON(result);
                    return Ok(jsonResult);
                }
            }
        }

        /// <summary>
        /// ////////////////////////////////////////////////////////////consluta 7//////////////////////////////////////////////////////////////
        /// </summary>
        /// <returns></returns>


        [HttpGet("Productos menos vendidos")]
        public IActionResult GetProducto_menos_vendido()
        {
            string connectionString = _configuration.GetConnectionString("OlapConnection");
            using (AdomdConnection connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT NON EMPTY { [Measures].[Cantidad Venta] } ON COLUMNS, NON EMPTY { BOTTOMCOUNT([Dimension Productos].[Nombre Producto].[Nombre Producto].ALLMEMBERS, 10, [Measures].[Cantidad Venta]) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM [Dimenciones Pan Plus] CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                using (AdomdCommand command = new AdomdCommand(query, connection))
                {
                    var result = command.ExecuteCellSet();
                    var jsonResult = TransformToJSON(result);
                    return Ok(jsonResult);
                }
            }
        }

        /// <summary>
        /// ////////////////////////////////////////////////////////////consluta 8//////////////////////////////////////////////////////////////
        /// </summary>
        /// <returns></returns>


        [HttpGet("Ventas de productos en el fin de semana")]
        public IActionResult GetVentas_prodcutos_el_fin_de_semana()
        {
            string connectionString = _configuration.GetConnectionString("OlapConnection");
            using (AdomdConnection connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT NON EMPTY {[Measures].[Cantidad Venta] } ON COLUMNS,  NON EMPTY { TOPCOUNT ( [Dimension Productos].[Nombre Producto].[Nombre Producto].ALLMEMBERS * [Dim Date].[Is Weekend].[True], 20  )} DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM [Dimenciones Pan Plus] CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                using (AdomdCommand command = new AdomdCommand(query, connection))
                {
                    var result = command.ExecuteCellSet();
                    var jsonResult = TransformToJSON(result);
                    return Ok(jsonResult);
                }
            }
        }

        /// <summary>
        /// ////////////////////////////////////////////////////////////consluta 9//////////////////////////////////////////////////////////////
        /// </summary>
        /// <returns></returns>


        [HttpGet("Ventas de productos por proveedor")]
        public IActionResult GetVentas_de_producto_por_proveedor()
        {
            string connectionString = _configuration.GetConnectionString("OlapConnection");
            using (AdomdConnection connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT NON EMPTY { [Measures].[Cantidad Venta] } ON COLUMNS, NON EMPTY { ([Dimension Productos].[Nombre Producto].[Nombre Producto].ALLMEMBERS * [Dimension Productos].[Nombre Proveedor].[Nombre Proveedor].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM [Dimenciones Pan Plus] CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                using (AdomdCommand command = new AdomdCommand(query, connection))
                {
                    var result = command.ExecuteCellSet();
                    var jsonResult = TransformToJSON(result);
                    return Ok(jsonResult);
                }
            }
        }


        /// <summary>
        /// ////////////////////////////////////////////////////////////consluta 9//////////////////////////////////////////////////////////////
        /// </summary>
        /// <returns></returns>


        [HttpGet("Top ingresos por cliente")]
        public IActionResult GetTop_ingresos_por_cliente()
        {
            string connectionString = _configuration.GetConnectionString("OlapConnection");
            using (AdomdConnection connection = new AdomdConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT NON EMPTY { [Measures].[Venta Total] } ON COLUMNS, NON EMPTY {TOPCOUNT(\r\n[Dimension Cliente].[Nombre Cliente].[Nombre Cliente].ALLMEMBERS, 20,         [Measures].[Venta Total])} DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME \r\nON ROWS FROM [Dimenciones Pan Plus] CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS";

                using (AdomdCommand command = new AdomdCommand(query, connection))
                {
                    var result = command.ExecuteCellSet();
                    var jsonResult = TransformToJSON(result);
                    return Ok(jsonResult);
                }
            }
        }


        private List<Dictionary<string, object>> TransformToJSON(CellSet result)
        {
            var jsonData = new List<Dictionary<string, object>>();
            int cellIndex = 0;

            foreach (var rowPosition in result.Axes[1].Positions)
            {
                var dataPoint = new Dictionary<string, object>();

                for (int i = 0; i < rowPosition.Members.Count; i++)
                {
                    var dimensionName = result.Axes[1].Set.Hierarchies[i].Name;
                    dataPoint[dimensionName] = rowPosition.Members[i].Caption;
                }

                for (int colIndex = 0; colIndex < result.Axes[0].Positions.Count; colIndex++)
                {
                    var measureName = result.Axes[0].Positions[colIndex].Members[0].Caption;
                    var cellValue = result.Cells[cellIndex].Value;

                    dataPoint[measureName] = cellValue;
                    cellIndex++;
                }

                jsonData.Add(dataPoint);
            }

            return jsonData;
        }
    }
}
