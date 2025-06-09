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
      console.log(Object.keys(row));
      console.log(row);
      const templateBytes = await fetch('facture_template.pdf').then(res => res.arrayBuffer());
      const pdfDoc = await PDFLib.PDFDocument.load(templateBytes);
      const helveticaFont = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);
      const page = pdfDoc.getPages()[0];
      const pageWidth = page.getWidth();

      // Fonctions utilitaires
      const rightAlign = (text, xRight, y, size = 11) => {
        const textWidth = helveticaFont.widthOfTextAtSize(String(text), size);
        page.drawText(String(text), {
          x: xRight - textWidth,
          y,
          size,
          font: helveticaFont,
        });
      };
      const draw = (text, x, y, size = 11) => {
        page.drawText(String(text), {
          x,
          y,
          size,
          font: helveticaFont,
        });
      };

      // Booking number (en haut à droite)
      rightAlign(row["Reservation number"], pageWidth - 40, 710, 12);

      // Date actuelle (en haut à droite, sous le numéro de réservation)
      const currentDate = new Date().toLocaleDateString('fr-FR', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
      });
      rightAlign(currentDate, pageWidth - 40, 800, 12);

      // Guest name (à gauche, sous "Guest information:")
      const guestName = (row["Guest name"] || row["Booker name"]).split(/[\r\n]+/)[0];
      draw(guestName, 40, 650, 12);

      // Total guests (en face de "Total guests:")
      const personsKey = Object.keys(row).find(
        k => k.trim().toLowerCase() === "persons"
      );
      draw(row[personsKey], 40, 605, 12);

      // Total units/rooms (en face de "Total units/rooms:")
      draw(row["Rooms"], 40, 655, 12);

      // Arrival (Check-in)
      rightAlign(row["Arrival"], pageWidth - 40, 660, 12);

      // Departure (Check-out)
      rightAlign(row["Departure"], pageWidth - 40, 620, 12);

      // Length of stay
      rightAlign(`${row["Room nights"]} night${parseInt(row["Room nights"]) > 1 ? "s" : ""}`, pageWidth - 40, 575, 12);

      // Approximate arrival time (toujours "No time provided")
      draw("No time provided", 40, 515, 12);

      // Total price (gauche, sous la ligne, en gras)
      const helveticaBold = await pdfDoc.embedFont(PDFLib.StandardFonts.HelveticaBold);
      page.drawText(`€ ${parseFloat(row["Final amount"]).toFixed(2)}`, {
        x: 40,
        y: 450,
        size: 12,
        font: helveticaBold,
      });

      // Commission amount(droite)
      rightAlign(`€ ${parseFloat(row["Commission amount"]).toFixed(2)}`, pageWidth - 40, 430, 12);

      // Commissionable amount (droite)
      // rightAlign(`€ ${parseFloat(row["Original amount"]).toFixed(2)}`, pageWidth - 70, 440, 12);

      // Détail séjour (ligne du bas, droite)
      rightAlign(`1 x € ${parseFloat(row["Original amount"]).toFixed(2)}`, pageWidth - 40, 460, 12);


      // Taxe de séjour (montant total, droite)
      // const taxeTotale = (2.5 * parseInt(row["Persons"]) || 0).toFixed(2);
      // rightAlign(`€ ${taxeTotale}`, pageWidth - 70, 480, 12);

      // Total unit/room price (gauche, bas)
      draw(`€ ${parseFloat(row["Final amount"]).toFixed(2)}`, pageWidth - 85, 260, 12);

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
