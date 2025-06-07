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

      // Fonctions utilitaires
      const draw = (text, x, y, size = 11) => {
        page.drawText(String(text), {
          x,
          y,
          size,
          font: helveticaFont,
        });
      };
      const rightAlign = (text, xRight, y, size = 11) => {
        const textWidth = helveticaFont.widthOfTextAtSize(String(text), size);
        page.drawText(String(text), {
          x: xRight - textWidth,
          y,
          size,
          font: helveticaFont,
        });
      };

      // --- En-tête à droite ---
      rightAlign("Booking number:", pageWidth - 50, 790, 11);
      rightAlign(row["Reservation number"], pageWidth - 50, 775, 11);

      // --- Infos à gauche ---
      draw("Guest information:", 50, 740, 11);
      draw(row["Guest name"], 50, 725, 11);
      draw(row["Country"], 50, 710, 11);
      draw("Total guests:", 50, 690, 11);
      draw(row["Persons"], 150, 690, 11);
      draw("Total units/rooms:", 50, 675, 11);
      draw(row["Rooms"], 150, 675, 11);
      draw("Preferred language:", 50, 660, 11);
      draw("Anglais", 150, 660, 11);
      draw("Approximate arrival time:", 50, 645, 11);
      draw("No time provided", 200, 645, 11);

      // --- Infos à droite ---
      rightAlign("Check-in:", pageWidth - 50, 740, 11);
      rightAlign(row["Arrival"], pageWidth - 50, 725, 11);
      rightAlign("Check-out:", pageWidth - 50, 710, 11);
      rightAlign(row["Departure"], pageWidth - 50, 695, 11);
      rightAlign("Length of stay:", pageWidth - 50, 680, 11);
      rightAlign(`${row["Room nights"]} night`, pageWidth - 50, 665, 11);

      // --- Total price à gauche, Commission à droite ---
      draw("Total price:", 50, 620, 11);
      draw(`€ ${parseFloat(row["Final amount"]).toFixed(2)}`, 120, 620, 11);
      rightAlign("Commission:", pageWidth - 50, 620, 11);
      rightAlign(`€ ${parseFloat(row["Commission amount"]).toFixed(2)}`, pageWidth - 50, 605, 11);
      rightAlign("Commissionable amount:", pageWidth - 50, 590, 11);
      rightAlign(`€ ${parseFloat(row["Original amount"]).toFixed(2)}`, pageWidth - 50, 575, 11);

      // --- Description chambre ---
      draw("Chambre Lits Jumeaux - Vue sur Piscine", 50, 540, 11);
      draw("Breakfast included", 50, 525, 11);

      // --- Détail séjour ---
      draw("02 - 03 Jun 2025", 50, 500, 11);
      draw("Non remboursable (TARIF STANDARD), " + row["Country"], 150, 500, 11);
      rightAlign(`1 x € ${parseFloat(row["Original amount"]).toFixed(2)}`, pageWidth - 50, 500, 11);

      // --- Taxes ---
      draw("Taxe de séjour", 50, 470, 11);
      draw("€ 2,50 par personne et par nuit", 50, 455, 11);
      const taxeTotale = (2.5 * parseInt(row["Persons"]) || 0).toFixed(2);
      draw(`€ ${taxeTotale}`, 150, 455, 11);

      // --- Total unit/room price ---
      draw("Total unit/room price", 50, 430, 11);
      draw(`€ ${parseFloat(row["Final amount"]).toFixed(2)}`, 200, 430, 11);

      // --- TVA ---
      draw("Rate includes", 50, 410, 9);
      draw("20 % de TVA", 120, 410, 9);

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
