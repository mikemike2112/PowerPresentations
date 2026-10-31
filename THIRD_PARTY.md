# Dipendenze di terze parti

Solo licenze MIT / Apache-2.0 / BSD. Nessun font, icona o risorsa Microsoft è incluso nel repository o nel binario:
i font si leggono a runtime dal sistema dell'utente, l'icona dell'app è originale (`tools/make-icon.ps1`).

## Prodotto (Core e App)

| Pacchetto | Versione | Licenza | Uso |
|---|---|---|---|
| DocumentFormat.OpenXml (+ .Framework) | 3.5.1 | MIT | lettura/scrittura OOXML |
| System.IO.Packaging | 10.0.2 | MIT | dipendenza di Open XML SDK |
| PDFsharp | **6.2.0** | MIT | esportazione PDF. Attenzione: dalla 6.3/6.5 la licenza di PDFsharp passa ad AGPL-3.0, quindi la versione è bloccata a 6.2.0 |
| Microsoft.Extensions.Logging / .Abstractions / DependencyInjection / Options / Primitives | 8.0.x | MIT | dipendenze transitive di PDFsharp |
| System.Security.Cryptography.Pkcs | 10.0.14 | MIT | dipendenza transitiva di PDFsharp (versione portata a una senza advisory note) |

## Solo test e strumenti

| Pacchetto | Versione | Licenza |
|---|---|---|
| xunit, xunit.runner.visualstudio | 2.9.3 / 3.1.4 | Apache-2.0 |
| Microsoft.NET.Test.Sdk | 17.14.1 | MIT |
| coverlet.collector | 6.0.4 | MIT |

`tools/PdfRaster` usa solo API di Windows (`Windows.Data.Pdf`) e WPF, nessun pacchetto.
