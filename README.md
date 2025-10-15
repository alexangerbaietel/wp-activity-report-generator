# Raport Activitate WordPress – Generator Automat `.docx`

Un script JavaScript care generează automat rapoarte lunare de activitate pentru autorii din WordPress, direct din consola browserului.  
Parcurge toate paginile din lista de articole, extrage linkurile publice și creează un fișier **Word (.docx)** bine formatat, numerotat și gata de utilizare.

---

## Ce face
- Parcurge automat paginile din `/wp-admin/edit.php` pentru un autor specific.
- Identifică articolele publicate într-o lună și un an introduse manual.
- Afișează un progress bar în consolă în timp real.
- Generează un fișier `.docx` curat, formatat cu font *Calibri*, titlu centrat și listă numerotată.
- Rulează complet local, fără pluginuri sau conectare externă.

---

## Cum se folosește
1. Autentifică-te în panoul de administrare WordPress.  
2. Accesează pagina cu articolele autorului.
3. Deschide consola browserului (`Ctrl + Shift + J` în Chrome).  
4. Lipește codul complet din fișierul `raport_activitate.js`.
5. !! IMPORTANT !! Schimbă XX din cod cu numărul autorului tău (poți găsi acest număr/ID în linkul de WordPress)
6. Apasă Enter și urmează instrucțiunile pentru lună și an.  
7. La final se va descărca automat un fișier `.docx` cu raportul.

---

## Tehnologii folosite
- **JavaScript (ES2022)** – rulat direct în browser  
- **docx.js** – generare fișiere `.docx`  
- **Fetch API + DOMParser** – colectare date din WordPress  

---

## Note
- Scriptul nu accesează date externe, nu modifică nimic în WordPress și nu folosește autentificare automată.  
- Rulează local, doar în sesiunea curentă a utilizatorului logat.  
- Testat în **Google Chrome** și compatibil cu WordPress 6.x.  

---

## Licență
MIT License – poți folosi, modifica și redistribui liber, cu menționarea autorului.
