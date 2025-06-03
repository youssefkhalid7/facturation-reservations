async function processExcel() {
  const fileInput = document.getElementById('excelFile');
  if (!fileInput.files[0]) return alert('Veuillez importer un fichier Excel.');

  const excelFile = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = async function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

    const JSZipInstance = new JSZip();

    for (const row of rows.filter(r => r.Status === 'OK')) {
      const templateBytes = await fetch('facture_template.pdf').then(res => res.arrayBuffer());
      const pdfDoc = await PDFLib.PDFDocument.load(templateBytes);
      const page = pdfDoc.getPages()[0];

      const draw = (text, x, y, size = 11) => {
        page.drawText(String(text), { x, y, size });
      };

      // Coordonnées basiques — à ajuster selon le modèle !
      draw(row["Reservation number"], 130, 715);
      draw(row["Guest name"], 130, 685);
      draw(row["Country"], 130, 670);
      draw(String(row["Persons"]), 130, 655);
      draw(String(row["Rooms"]), 130, 640);
      draw("Aucune heure fournie", 180, 625);
      draw(row["Arrival"], 130, 595);
      draw(row["Departure"], 130, 580);
      draw(`${row["Room nights"]} night`, 130, 565);
      draw(`€ ${row["Final amount"]}`, 130, 535);
      draw(`€ ${row["Commission amount"]}`, 130, 505);
      draw(`€ ${row["Original amount"]}`, 200, 490);

      const taxeTotale = (2.5 * parseInt(row["Persons"]) || 0).toFixed(2);
      draw(`€ ${taxeTotale}`, 130, 430);

      draw(`€ ${row["Final amount"]}`, 180, 405); // Total unit/room price

      const pdfBytes = await pdfDoc.save();
      const safeName = `${(row["Booker name"] || "facture").replace(/[^a-z0-9]/gi, '_')}_${row["Reservation number"]}.pdf`;
      JSZipInstance.file(safeName, pdfBytes);
    }

    const zipBlob = await JSZipInstance.generateAsync({ type: 'blob' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(zipBlob);
    link.download = 'factures_clients.zip';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  reader.readAsArrayBuffer(excelFile);
}
