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

      const draw = (text, x, y, size = 11) => {
        page.drawText(String(text), {
          x,
          y,
          size,
          font: helveticaFont,
        });
      };

      // Commencer en haut à gauche et descendre ligne par ligne
      draw(`Booking number: ${row["Reservation number"]}`, 50, 690);
      draw(`Guest information: ${row["Guest name"]}`, 50, 670);
      draw(`${row["Country"]}`, 50, 655);

      draw(`Total guests: ${row["Persons"]}`, 50, 630);
      draw(`Total units/rooms: ${row["Rooms"]}`, 50, 615);
      draw(`Approximate arrival time: No time provided`, 50, 600);

      draw(`Check-in: ${row["Arrival"]}`, 50, 575);
      draw(`Check-out: ${row["Departure"]}`, 50, 560);
      draw(`Length of stay: ${row["Room nights"]} night`, 50, 545);

      draw(`Total price: € ${parseFloat(row["Final amount"]).toFixed(2)}`, 50, 520);
      draw(`Commission: € ${parseFloat(row["Commission amount"]).toFixed(2)}`, 50, 505);
      draw(`Commissionable amount: € ${parseFloat(row["Original amount"]).toFixed(2)}`, 50, 490);

      // Description de la chambre
      draw(`Chambre Lits Jumeaux - Vue sur Piscine`, 50, 465);
      draw(`Breakfast included`, 50, 450);
      draw(`02 - 03 Jun 2025 Non remboursable (TARIF STANDARD), ${row["Country"]}`, 50, 435);
      draw(`1 x € ${parseFloat(row["Original amount"]).toFixed(2)}`, 50, 420);

      // Taxes
      draw(`Taxe de séjour`, 50, 395);
      draw(`€ 2,50 par personne et par nuit`, 50, 380);

      const taxeTotale = (2.5 * parseInt(row["Persons"]) || 0).toFixed(2);
      draw(`€ ${taxeTotale}`, 50, 365);

      draw(`Total unit/room price € ${parseFloat(row["Final amount"]).toFixed(2)}`, 50, 340);
      draw(`Rate includes 20 % de TVA`, 50, 325);

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
