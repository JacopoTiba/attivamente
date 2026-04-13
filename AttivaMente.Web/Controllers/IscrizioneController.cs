using AttivaMente.Core.Models;
using AttivaMente.Core.OfficeAutomation;
using AttivaMente.Data;
using AttivaMente.Web.ViewModels;
using Microsoft.AspNetCore.Mvc;

namespace AttivaMente.Web.Controllers
{
    public class IscrizioneController : Controller
    {
        private readonly IscrizioneRepository _repoIscrizioni;
        private readonly UtenteRepository _repoUtenti;

        public IscrizioneController(IConfiguration configuration)
        {
            string connStr = configuration.GetConnectionString("DefaultConnection");
            _repoIscrizioni = new IscrizioneRepository(connStr);
            _repoUtenti = new UtenteRepository(connStr);
        }

        public IActionResult Index(int? anno, bool isSoloIscritti = true, string? ricerca = null)
        {
            int currentYear = DateTime.Now.Year;
            var years = _repoIscrizioni.GetYears();

            if (!years.Contains(currentYear))
                years.Insert(0, currentYear);

            // Questa riga sistema eventuali incongruenze a db (iscrizioni in anni futuri)
            years = years.Distinct().OrderByDescending(x => x).ToList();

            int selectedYear = anno ?? currentYear;

            var model = new IscrizioniIndexViewModel
            {
                SelectedYear = selectedYear,
                IsSoloIscritti = isSoloIscritti,
                Ricerca = ricerca ?? "",
                Years = years
            };

            if (isSoloIscritti)
            {
                var iscritti = _repoIscrizioni.GetByYear(selectedYear);
                foreach (var item in iscritti)
                {
                    model.Rows.Add(new IscrizioneRowViewModel
                    {
                        UtenteId = item.UtenteId,
                        Cognome = item.Utente!.Cognome,
                        Nome = item.Utente!.Nome,
                        Email = item.Utente!.Email,
                        Tipo = item.Tipo,
                        DataIscrizione = item.DataIscrizione,
                        Azione = "Cancella"
                    });
                }
            }
            else
            {
                var utenti = _repoUtenti.GetAll();
                var iscrizioniAnno = _repoIscrizioni.GetByYear(selectedYear);
                foreach (var item in utenti)
                {
                    var iscrizioneAnnoCorrente = iscrizioniAnno.FirstOrDefault(i => i.UtenteId == item.Id);
                    bool isIscrittoAnnoSelezionato = iscrizioneAnnoCorrente != null;
                    bool isIscrittoAnnoPrecedente = _repoIscrizioni.Exists(item.Id, selectedYear - 1);

                    string azione;
                    string? tipo = null;
                    DateTime? dataIscrizione = null;

                    if (isIscrittoAnnoSelezionato)
                    {
                        azione = "Cancella";
                        tipo = iscrizioneAnnoCorrente!.Tipo;
                        dataIscrizione = iscrizioneAnnoCorrente.DataIscrizione;
                    }
                    else if (isIscrittoAnnoPrecedente)
                        azione = "Rinnova";
                    else
                        azione = "Iscrivi";

                    model.Rows.Add(new IscrizioneRowViewModel
                    {
                        UtenteId = item.Id,
                        Cognome = item.Cognome,
                        Nome = item.Nome,
                        Email = item.Email,
                        Tipo = tipo,
                        DataIscrizione = dataIscrizione,
                        Azione = azione
                    });
                }
            }

            if (!string.IsNullOrWhiteSpace(ricerca))
            {
                string testo = ricerca.Trim().ToLower();

                model.Rows = model.Rows
                    .Where(r =>
                        r.Nome.ToLower().Contains(testo) ||
                        r.Cognome.ToLower().Contains(testo) ||
                        $"{r.Cognome} {r.Nome}".ToLower().Contains(testo) ||
                        $"{r.Nome} {r.Cognome}".ToLower().Contains(testo))
                    .ToList();
            }

            return View(model);
        }

        [HttpPost]
        public IActionResult Delete(int utenteId, int anno, bool isSoloIscritti, string? ricerca = null,bool fromDettaglio=false)
        {
            _repoIscrizioni.Delete(utenteId, anno);
            if (fromDettaglio)
                return RedirectToAction("Details", "Utente", new { id = utenteId });

            return RedirectToAction("Index", new { anno, isSoloIscritti, ricerca });
        }
        //compito
        [HttpPost]
        public IActionResult Update([FromBody] UpdateIscrizioneRequest request)
        {
            int result = _repoIscrizioni.Update(request.UtenteId, request.Anno, request.Tipo, DateTime.Parse(request.DataIscrizione));
            if (result > 0)
                return Ok();
            else
                return BadRequest();
        }

        public IActionResult Storico(int utenteId)
        {
            var utente = _repoUtenti.GetById(utenteId);
            if (utente == null) return NotFound();
            ViewBag.Utente = utente;
            var iscrizioni = _repoIscrizioni.GetByUtente(utenteId);
            return View(iscrizioni);
        }
        public IActionResult Statistiche()
        {
            var statistiche = _repoIscrizioni.GetStatistiche();
            return View(statistiche);
        }

        [HttpPost]
        public IActionResult Renew(int utenteId, int anno, bool isSoloIscritti, string? ricerca = null,bool fromDettaglio=false)
        {
            if (!_repoIscrizioni.Exists(utenteId, anno))
            {
                _repoIscrizioni.Insert(utenteId, anno, "Rinnovo");
            }
            if (fromDettaglio)
                return RedirectToAction("Details", "Utente", new { id = utenteId });

            return RedirectToAction("Index", new { anno, isSoloIscritti, ricerca });
        }

        [HttpPost]
        public IActionResult Subscribe(int utenteId, int anno, bool isSoloIscritti, string? ricerca = null, bool fromDettaglio = false)
        {
            if (!_repoIscrizioni.Exists(utenteId, anno))
                _repoIscrizioni.Insert(utenteId, anno, "Nuova");

            if (fromDettaglio)
                return RedirectToAction("Details", "Utente", new { id = utenteId });

            return RedirectToAction("Index", new { anno, isSoloIscritti, ricerca });
        }



        public IActionResult CreateXlsx(int? anno)
        {
            var iscrizioni = anno.HasValue
                ? _repoIscrizioni.GetByYear(anno.Value)
                : _repoIscrizioni.GetAll(); // ← tutti gli anni

            byte[] fileBytes = ExcelAutomation.GetIscrizioniXlsxBytes(iscrizioni);
            return File(
                fileBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"Iscrizioni_{anno?.ToString() ?? "tutti"}.xlsx");
        }
    }
}
public class UpdateIscrizioneRequest //DTO per asp net
{
    public int UtenteId { get; set; }
    public int Anno { get; set; }
    public string Tipo { get; set; } = "";
    public string DataIscrizione { get; set; } = "";
}