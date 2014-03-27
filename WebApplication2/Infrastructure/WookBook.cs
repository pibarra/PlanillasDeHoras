using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebApplication2.Infrastructure
{
    public class WookBook
    {
        public static XLWorkbook GenerarWookBook() 
        {
            var Dias = new Dictionary<int, bool>();
            int CantDias = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month);
            for (int i = 1; i <= CantDias; i++)
            {
                Dias.Add(i, i % 6 == 0 ? true : false);
            }
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Sheet 1");
            worksheet.Range("A1:I1").Merge()
                .SetValue("PLANILLA MANUAL DE HORARIOS DE ENTRADA Y SALIDA")
                .Style.Font.SetBold(true)
                      .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                      .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                
            worksheet.Range("A2:I2").Merge()
                .SetValue("SECRETARIA DE ESTADO DE TRABAJO-RESOLUCION GENERAL 172/00")
                .Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                      .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            foreach (var day in Dias)
            {
                var NoLaboral = Convert.ToDateTime(day.Key + "/03/2014");
                Boolean feriado = NoLaboral.DayOfWeek == DayOfWeek.Saturday || NoLaboral.DayOfWeek == DayOfWeek.Sunday ? true : false;
                string celda = "A" + (day.Key + 2);
                string valor = day.Key.ToString();
                var color = feriado ? XLColor.Gray : XLColor.NoColor;
                worksheet.Cell(celda).SetValue(valor);
                
                worksheet.Range(celda +":I" +(day.Key + 2)).Style.Fill.SetBackgroundColor(color);
                
            }
            int lastday = Dias.OrderByDescending(x => x.Key).FirstOrDefault().Key;
            worksheet.Range("A3:I" + (lastday + 2))
                .Style.Border.SetInsideBorder(XLBorderStyleValues.Thin)
                .Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            return workbook;
        }
    }
}