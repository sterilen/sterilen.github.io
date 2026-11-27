# Lagersök – Sterilförråd, Södersjukhuset

Webbsidan för sökning av artiklar och platser i sterilförrådet (lager 3350), publicerad med GitHub Pages.

| Fil | Vad den är |
|---|---|
| `index.html` | Söksidan Lagersök |
| `statistik.html` | Statistik över sökningar och etikettutskrifter |
| `stad.html` | Städning med QR-skanning |
| `data.js` | Artikeldatabasen. Skapas av Lagersök Admin, ändra inte för hand |
| `status.json` | Underlag till Översikt: databasstatus, avvikelser och ändringar. Skapas av Lagersök Admin |
| `bilder/` | Artikelbilder (800 px och 200 px tumnagel, webp). Hanteras under BILDER i Lagersök Admin |
| `alternativa_nummer.csv` | Alternativa artikelnummer (nummer på låda eller hylla som inte finns i databasen) |
| `lagerid.txt` | Beställare till 3350 (deras skåplistor läses vid databasuppdateringen) |
| `renlager.txt` | Lager id för renförrådet |
| `sw.js`, `manifest.webmanifest`, ikoner | Gör att sidan kan installeras och fungera utan nät |
| `html5-qrcode.min.js` | QR-skannern till städsidan ([html5-qrcode](https://github.com/mebjas/html5-qrcode) 2.3.8, Apache 2.0) |

Programmet Lagersök Admin och dess källkod ligger i ett separat, privat repo.
