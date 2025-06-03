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

      // Match layout from reference invoice (4490222232.pdf)
      draw(row["Reservation number"], 120, 725); // Booking number
      draw(row["Guest name"], 120, 700);         // Guest name
      draw(row["Country"], 120, 685);            // Country
      draw(String(row["Persons"]), 120, 660);    // Total guests
      draw(String(row["Rooms"]), 120, 645);      // Total units/rooms
      draw("No time provided", 190, 630);        // Approx arrival time

      draw(row["Arrival"], 120, 605);            // Check-in
      draw(row["Departure"], 120, 590);          // Check-out
      draw(`${row["Room nights"]} night`, 120, 570); // Length of stay

      draw(`€ ${parseFloat(row["Final amount"]).toFixed(2)}`, 120, 545); // Total price
      draw(`€ ${parseFloat(row["Commission amount"]).toFixed(2)}`, 120, 515); // Commission
      draw(`€ ${parseFloat(row["Original amount"]).toFixed(2)}`, 190, 500); // Commissionable amount

      const taxeTotale = (2.5 * parseInt(row["Persons"]) || 0).toFixed(2);
      draw(`€ ${taxeTotale}`, 120, 445); // Tourist tax
      draw(`€ ${parseFloat(row["Final amount"]).toFixed(2)}`, 190, 425); // Total unit/room price

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
