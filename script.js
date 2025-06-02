async function processExcel() {
  const fileInput = document.getElementById('excelFile');
  if (!fileInput.files[0]) return alert('Please upload an Excel file.');

  const file = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = async function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

    const JSZipInstance = new JSZip();
    const { jsPDF } = window.jspdf;

    const logoDataUrl = await fetch('logo.webp')
      .then(res => res.blob())
      .then(blob => new Promise(resolve => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.readAsDataURL(blob);
      }));

    const okReservations = rows.filter(r => r.Status === 'OK');

    for (const row of okReservations) {
      const doc = new jsPDF({ unit: 'pt', format: 'a4' });
      let y = 50;

      // Logo
      doc.addImage(logoDataUrl, 'PNG', 40, y, 150, 50);
      y += 80;

      doc.setFontSize(10);
      doc.setTextColor(150); // Light gray for Booking number
      doc.text('Booking number:', 400, 50);
      doc.setTextColor(0);
      doc.text(row["Reservation number"].toString(), 500, 50);

      // Draw a separator line
      y += 10;
      doc.setDrawColor(200); // Light gray separator
      doc.setLineWidth(0.5);
      doc.line(40, y, 550, y);
      y += 15;

      // Guest Information section
      doc.setFontSize(12);
      doc.setTextColor(100);
      doc.text('Guest information:', 40, y);
      y += 15;

      doc.setFontSize(11);
      doc.setTextColor(0);
      doc.text(row["Guest name"], 40, y); y += 15;
      doc.text(row["Country"], 40, y); y += 15;
      doc.text(`Total guests: ${row["Persons"]}`, 40, y); y += 15;
      doc.text(`Total units/rooms: ${row["Rooms"]}`, 40, y); y += 15;
      doc.text(`Preferred language: English`, 40, y); y += 15;
      doc.text(`Approximate arrival time: No time provided`, 40, y); y += 20;

      // Draw separator
      doc.setDrawColor(200);
      doc.line(40, y, 550, y);
      y += 15;

      // Check-in/out block
      doc.setFontSize(10);
      doc.setTextColor(150);
      doc.text('Check-in:', 400, y - 90);
      doc.setTextColor(0);
      doc.text(row["Arrival"], 470, y - 90);
      doc.setTextColor(150);
      doc.text('Check-out:', 400, y - 75);
      doc.setTextColor(0);
      doc.text(row["Departure"], 470, y - 75);
      doc.setTextColor(150);
      doc.text('Length of stay:', 400, y - 60);
      doc.setTextColor(0);
      doc.text(`${row["Room nights"]} night(s)`, 500, y - 60);

      // Total price and commission block
      doc.setFontSize(12);
      doc.text('Total price:', 40, y);
      doc.setFontSize(11);
      doc.text(`€ ${row["Final amount"]}`, 120, y);
      y += 30;

      doc.setFontSize(12);
      doc.text('Commission:', 400, y - 30);
      doc.setFontSize(11);
      doc.text(`€ ${row["Commission amount"]}`, 480, y - 30);
      doc.setFontSize(12);
      doc.text('Commissionable amount:', 400, y - 15);
      doc.setFontSize(11);
      doc.text(`€ ${row["Original amount"]}`, 530, y - 15);

      // Draw separator
      y += 10;
      doc.setDrawColor(200);
      doc.line(40, y, 550, y);
      y += 15;

      // Booking summary with proper styling
      doc.setFontSize(12);
      doc.setTextColor(0);
      doc.text('Twin Beds Room - Pool View', 40, y); y += 15;
      doc.text('Breakfast included', 40, y); y += 15;

      doc.setFontSize(11);
      const detailText = '02 - 03 Jun 2025 Non-refundable (STANDARD RATE), United States (Non-refundable (STANDARD RATE) -10%) 1 x € 68,85';
      doc.text(detailText, 40, y, { maxWidth: 500 }); y += 30;

      doc.setFontSize(11);
      doc.text(`Total unit/room price € ${row["Final amount"]}`, 40, y); y += 15;
      doc.text('Rate includes 20% VAT', 40, y);

      // Final separator
      y += 20;
      doc.setDrawColor(200);
      doc.line(40, y, 550, y);

      // Save PDF to ZIP
      const pdfBlob = doc.output('blob');
      const buffer = await pdfBlob.arrayBuffer();
      const safeName = `${(row["Booker name"] || "invoice").replace(/[^a-z0-9]/gi, '_')}_${row["Reservation number"] || Date.now()}.pdf`;
      JSZipInstance.file(safeName, buffer);
    }

    const zipBlob = await JSZipInstance.generateAsync({ type: 'blob' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(zipBlob);
    link.download = 'customer_invoices.zip';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  reader.readAsArrayBuffer(file);
}
