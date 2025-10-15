(async () => {
  
  const mod = await import("https://cdn.jsdelivr.net/npm/docx@9.1.0/+esm");
  const { Document, Packer, Paragraph, TextRun, AlignmentType } = mod;

  // === Prompturi luna/an ===
  const monthInput = prompt("Introdu luna (1-12):");
  if (!monthInput) return;
  const yearInput = prompt("Introdu anul (ex: 2025):");
  if (!yearInput) return;
  const month = parseInt(monthInput, 10), year = parseInt(yearInput, 10);
  if (isNaN(month) || isNaN(year)) return;

  // === Config ===
  const baseOrigin = window.location.origin;
  const basePath = "/wp-admin/edit.php";
  const paramsBase = new URLSearchParams({ author: "63", post_type: "post" });

  const parseDate = (t) => {
    const m = (t || "").match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);
    return m ? { d: +m[1], m: +m[2], y: +m[3] } : null;
  };

  const fetchPage = async (p) => {
    const url = `${baseOrigin}${basePath}?${paramsBase}&paged=${p}`;
    const r = await fetch(url, { credentials: "same-origin" });
    const text = await r.text();
    return new DOMParser().parseFromString(text, "text/html");
  };

  // === Functie progress bar ===
  function showProgress(current, totalEstimate = 30) {
    const percent = Math.min(100, Math.round((current / totalEstimate) * 100));
    const totalBars = 25;
    const filledBars = Math.min(totalBars, Math.round((percent / 100) * totalBars));
    const emptyBars = totalBars - filledBars;
    const bar = "█".repeat(filledBars) + "░".repeat(emptyBars);
    console.clear();
    console.log(`Extragem articolele din luna ${month}.${year}...`);
    console.log(`[${bar}] ${percent}% complet`);
  }

  // === Colectare articole ===
  const posts = [];
  let page = 1;
  let foundPrev = false;
  let maxPagesEstimate = 30; // estimare initiala (ajustabila)
  console.log("Incep colectarea articolelor...");

  while (true) {
    showProgress(page, maxPagesEstimate);
    const doc = await fetchPage(page);
    const rows = [...doc.querySelectorAll("tbody#the-list tr")];
    if (!rows.length) break;

    let found = false;
    for (const r of rows) {
      const a = r.querySelector('a[aria-label^="Vezi"]');
      const dcell = r.querySelector("td.date.column-date");
      if (!a || !dcell) continue;
      const dt = parseDate(dcell.innerText);
      if (dt && dt.m === month && dt.y === year) {
        posts.push({ d: dt.d, link: a.href.trim() });
        found = true;
      }
    }

    if (found) foundPrev = true;
    else if (foundPrev && !found) break;

    page++;
    if (page > 2000) break;
    if (page > maxPagesEstimate) maxPagesEstimate = page + 5; // extind estimarea daca se continua
    await new Promise(r => setTimeout(r, 200));
  }

  console.clear();
  console.log("Colectarea s-a Incheiat.");
  console.log(`Total articole gasite: ${posts.length}`);

  if (!posts.length) {
    alert("Nu s-au gasit articole pentru luna selectata.");
    return;
  }

  // === Curatare / sortare ===
  const seen = new Set();
  const uniq = posts.filter(p => !seen.has(p.link) && seen.add(p.link))
                    .sort((a, b) => a.d - b.d);

  // === Creare document Word ===
  const monthNames = [
    "Ianuarie","Februarie","Martie","Aprilie","Mai","Iunie",
    "Iulie","August","Septembrie","Octombrie","Noiembrie","Decembrie"
  ];
  const niceMonth = monthNames[month - 1] || `Luna ${month}`;
  const titleText = `RAPORT DE ACTIVITATE - ${niceMonth.toUpperCase()} ${year}`;

  const children = [];

  // Titlu 
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 500 },
      children: [
        new TextRun({
          text: titleText,
          bold: true,
          size: 36,
          font: "Calibri"
        })
      ]
    })
  );

  // Lista ordonata
  uniq.forEach((it, i) => {
    children.push(
      new Paragraph({
        spacing: { after: 120 },
        children: [
          new TextRun({
            text: `${i + 1}. `,
            bold: true,
            size: 24,
            font: "Calibri"
          }),
          new TextRun({
            text: it.link,
            size: 24,
            font: "Calibri"
          })
        ]
      })
    );
  });

  const docx = new Document({
    sections: [{ properties: {}, children }]
  });

  const blob = await Packer.toBlob(docx);
  const filename = `Raport Activitate - ${niceMonth} ${year}.docx`;

  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);

  console.clear();
  console.log(`Raport generat: ${filename}`);
  console.log(`Total articole: ${uniq.length}`);
})();
