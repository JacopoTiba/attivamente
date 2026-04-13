using AttivaMente.Core.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttivaMente.Core.OfficeAutomation
{
    public static class ExcelAutomation
    {
        public static byte[] GetUsersXlsxBytes(List<Utente> utenti)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                // costruisco i contenuti del file
                var ws = package.Workbook.Worksheets.Add("Utenti");

                // intestazioni
                ws.Cells[1, 1].Value = "Id";
                ws.Cells[1, 2].Value = "Nome";
                ws.Cells[1, 3].Value = "Cognome";
                ws.Cells[1, 4].Value = "Email";
                ws.Cells[1, 5].Value = "Ruolo";
                // dati
                int row = 2;
                foreach (var utente in utenti)
                {
                    ws.Cells[row, 1].Value = utente.Id;
                    ws.Cells[row, 2].Value = utente.Nome;
                    ws.Cells[row, 3].Value = utente.Cognome;
                    ws.Cells[row, 4].Value = utente.Email;
                    ws.Cells[row, 5].Value = utente.Ruolo!.Nome;
                    row++;
                }
                // larghezza colonne automatica
                ws.Cells.AutoFitColumns();

                return package.GetAsByteArray();
            }
        }

        public static byte[] GetIscrizioniXlsxBytes(List<Iscrizione> Iscrizioni)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                // costruisco i contenuti del file
                var ws = package.Workbook.Worksheets.Add("Iscrizioni");

                // intestazioni
                ws.Cells[1, 1].Value = "Id";
                ws.Cells[1, 2].Value = "Nome";
                ws.Cells[1, 3].Value = "Cognome";
                ws.Cells[1, 4].Value = "Anno";
                ws.Cells[1, 5].Value = "Tipo";
                ws.Cells[1, 6].Value = "DataIscrizione";

                // dati
                int row = 2;
                foreach (var iscrizione in Iscrizioni)
                {
                    ws.Cells[row, 1].Value = iscrizione.UtenteId;
                    ws.Cells[row, 2].Value = iscrizione.Utente!.Nome;
                    ws.Cells[row, 3].Value = iscrizione.Utente!.Cognome;
                    ws.Cells[row, 4].Value = iscrizione.Anno;
                    ws.Cells[row, 5].Value = iscrizione.Tipo;
                    ws.Cells[row, 6].Value = iscrizione.DataIscrizione;
                    ws.Cells[row, 6].Style.Numberformat.Format = "dd/MM/yyyy";
                    row++;
                }
                // larghezza colonne automatica
                ws.Cells.AutoFitColumns();

                return package.GetAsByteArray();
            }
        }
    }
}
