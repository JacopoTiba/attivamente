using AttivaMente.Core.Models;
using AttivaMente.Data;
using Microsoft.AspNetCore.Mvc;

namespace AttivaMente.Web.Controllers.Api
{
    [ApiController] // indispensabile
    [Route("api/utenti")] // prefisso che vale per tutti gli endpoint del controller
    public class UtenteApiController : ControllerBase
    {
        private readonly UtenteRepository _repoUtenti;
        public UtenteApiController(IConfiguration configuration)
        {
            string connStr = configuration.GetConnectionString("DefaultConnection");
            _repoUtenti = new UtenteRepository(connStr);
        }

        //[Route("all")] // suffisso route specifica di questo endpoint
        [HttpGet("all")] //sa dove e come trovarli
        public IActionResult Index()
        {
            var utenti = _repoUtenti.GetAll();
            return utenti.Count == 0 ? NotFound() : Ok(utenti);
        }

        //[Route("{id}")]
        [HttpGet("{id}")] // suffisso di route che esplicita il verbo HttpGet (chiamata in GET)
        public IActionResult Detail(int id)
        {
            var utente = _repoUtenti.GetById(id);
            return Ok(utente);
        }

        [HttpGet("search")] // esempio parametrizzazione in querystring
        public IActionResult Search(
            string pattern,
            int? ruolo,
            string? order,
            string? direction)
        {
            // tutti i parametri (alcuni facoltativi) sono automaticamente passabili in querystring
            // https://localhost:7137/api/utenti/search?pattern=al&ruolo=2&order=Nome&direction=DESC
            var utenti = _repoUtenti.Search(pattern, ruolo, order, direction);
            return utenti.Count == 0 ? NotFound() : Ok(utenti);
        }

        [HttpPost("add")]
        public IActionResult Add([FromBody] Utente utente)
        {
            _repoUtenti.Add(utente);
            return Ok();
        }

        [HttpPut("update")]
        public IActionResult Update([FromBody] Utente utente)
        {
            _repoUtenti.Add(utente);
            return Ok();
        }

        [HttpDelete("delete/{id}")]
        public IActionResult Delete(int id)
        {
            _repoUtenti.Delete(id);
            return Ok();
        }
    }
}
