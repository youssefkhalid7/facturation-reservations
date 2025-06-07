async function processExcel() {
  const fileInput = document.getElementById('excelFile');
  if (!fileInput.files[0]) return alert('Please upload an Excel file.');

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
      const helveticaFont = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);
      const page = pdfDoc.getPages()[0];
      const pageWidth = page.getWidth();

      // Fonction pour aligner à droite
      const rightAlign = (text, xRight, y, size = 11) => {
        const textWidth = helveticaFont.widthOfTextAtSize(String(text), size);
        page.drawText(String(text), {
          x: xRight - textWidth,
          y,
          size,
          font: helveticaFont,
        });
      };
      // Fonction pour écrire à gauche
      const draw = (text, x, y, size = 11) => {
        page.drawText(String(text), {
          x,
          y,
          size,
          font: helveticaFont,
        });
      };

      // --- Remplir uniquement les valeurs, coordonnées à ajuster selon le template ---
      // Booking number (en haut à droite)
      rightAlign(row["Reservation number"], pageWidth - 50, 775, 11);

      // Guest name (à gauche)
      draw(row["Guest name"], 50, 725, 11);

      // Country (sous le nom)
      draw(row["Country"], 50, 710, 11);

      // Total guests
      draw(row["Persons"], 150, 690, 11);

      // Total units/rooms
      draw(row["Rooms"], 150, 675, 11);

      // Arrival (Check-in)
      rightAlign(row["Arrival"], pageWidth - 50, 725, 11);

      // Departure (Check-out)
      rightAlign(row["Departure"], pageWidth - 50, 695, 11);

      // Length of stay
      rightAlign(`${row["Room nights"]} night`, pageWidth - 50, 665, 11);

      // Total price
      draw(`€ ${parseFloat(row["Final amount"]).toFixed(2)}`, 120, 620, 11);

      // Commission
      rightAlign(`€ ${parseFloat(row["Commission amount"]).toFixed(2)}`, pageWidth - 50, 605, 11);

      // Commissionable amount
      rightAlign(`€ ${parseFloat(row["Original amount"]).toFixed(2)}`, pageWidth - 50, 575, 11);

      // Détail séjour (ligne du bas)
      rightAlign(`1 x € ${parseFloat(row["Original amount"]).toFixed(2)}`, pageWidth - 50, 500, 11);

      // Taxe de séjour (montant total)
      const taxeTotale = (2.5 * parseInt(row["Persons"]) || 0).toFixed(2);
      draw(`€ ${taxeTotale}`, 150, 455, 11);

      // Total unit/room price
      draw(`€ ${parseFloat(row["Final amount"]).toFixed(2)}`, 200, 430, 11);

      // --- Génération du PDF ---
      const pdfBytes = await pdfDoc.save();
      const safeName = `${(row["Booker name"] || "invoice").replace(/[^a-z0-9]/gi, '_')}_${row["Reservation number"]}.pdf`;
      JSZipInstance.file(safeName, pdfBytes);
    }

    const zipBlob = await JSZipInstance.generateAsync({ type: 'blob' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(zipBlob);
    link.download = 'client_invoices.zip';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  reader.readAsArrayBuffer(excelFile);
}
