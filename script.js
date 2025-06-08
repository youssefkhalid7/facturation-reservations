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

      // Booking number (sous "Booking number:")
      rightAlign(row["Reservation number"], pageWidth - 60, 740, 12);

      // Guest name (sous "Guest information:")
      draw(row["Guest name"], 70, 705, 12);

      // Total guests (sous "Total guests:")
      draw(row["Persons"], 70, 675, 12);

      // Total units/rooms (sous "Total units/rooms:")
      draw(row["Rooms"], 70, 650, 12);

      // Arrival (sous "Check-in:")
      rightAlign(row["Arrival"], pageWidth - 70, 705, 12);

      // Departure (sous "Check-out:")
      rightAlign(row["Departure"], pageWidth - 70, 675, 12);

      // Length of stay (sous "Length of stay:")
      rightAlign(`${row["Room nights"]} night${parseInt(row["Room nights"]) > 1 ? "s" : ""}`, pageWidth - 70, 650, 12);

      // Total price (sous "Total price:")
      draw(`€ ${parseFloat(row["Final amount"]).toFixed(2)}`, 70, 600, 12);

      // Commission (sous "Commission:")
      rightAlign(`€ ${parseFloat(row["Commission amount"]).toFixed(2)}`, pageWidth - 70, 600, 12);

      // Commissionable amount (sous "Commissionable amount:")
      rightAlign(`€ ${parseFloat(row["Original amount"]).toFixed(2)}`, pageWidth - 70, 580, 12);

      // Détail séjour (sous la description, à droite)
      rightAlign(`1 x € ${parseFloat(row["Original amount"]).toFixed(2)}`, pageWidth - 70, 510, 12);

      // Taxe de séjour (montant total, droite)
      const taxeTotale = (2.5 * parseInt(row["Persons"]) || 0).toFixed(2);
      rightAlign(`€ ${taxeTotale}`, pageWidth - 70, 480, 12);

      // Total unit/room price (sous "Total unit/room price")
      draw(`€ ${parseFloat(row["Final amount"]).toFixed(2)}`, 170, 450, 12);

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
