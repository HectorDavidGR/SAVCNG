using System;

namespace SAVCNG_ExcelDNA.Core
{
    // =========================================================================
    // DTO: OBJETO DE TRANSFERENCIA DE RESULTADOS (MVP)
    // =========================================================================
    // Mentoría: Este objeto permite que la lógica de negocio se comunique con
    // la capa de presentación (UI) sin depender de controles de Windows Forms.
    public class ResultadoValidacion
    {
        public bool Exito { get; set; }
        public string Mensaje { get; set; }
        public bool AlertaInyectada { get; set; }

        // Constructor por defecto para inicializar estados seguros
        public ResultadoValidacion()
        {
            Exito = false;
            Mensaje = string.Empty;
            AlertaInyectada = false;
        }
    }
}