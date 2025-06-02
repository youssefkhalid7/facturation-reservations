async function processExcel() {
  const fileInput = document.getElementById('excelFile');
  if (!fileInput.files[0]) return alert('Please upload an Excel file.');

  const file = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = async function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);

    const JSZipInstance = new JSZip();
    const { jsPDF } = window.jspdf;

    // Load logo
    const logoDataUrl = await fetch('logo.webp')
      .then(res => res.blob())
      .then(blob => new Promise(resolve => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.readAsDataURL(blob);
      }));

    const okReservations = rows.filter(r => r.Status === 'OK');

    for (const row of okReservations) {
      const doc = new jsPDF();

      // Logo and Title
      doc.addImage(logoDataUrl, 'PNG', 10, 10, 40, 20);
      doc.setFontSize(18);
      doc.text('CUSTOMER INVOICE', 70, 20);

      doc.setFontSize(11);
      let y = 40;

      doc.text(`Booking number: ${row["Reservation number"]}`, 20, y); y += 10;
      doc.text(`Guest name: ${row["Booker name"]}`, 20, y); y += 10;
      doc.text(`Country: ${row["Country"]}`, 20, y); y += 10;
      doc.text(`Check-in: ${row["Arrival"]}`, 20, y);
      doc.text(`Check-out: ${row["Departure"]}`, 120, y); y += 10;
      doc.text(`Length of stay: ${row["Room nights"]} night(s)`, 20, y); y += 10;
      doc.text(`Room(s): ${row["Rooms"]}`, 20, y);
      doc.text(`Guests: ${row["Persons"]}`, 120, y); y += 10;

      // Financial Details
      y += 5;
      doc.setFontSize(12);
      doc.text('--- Billing Details ---', 20, y); y += 10;
      doc.setFontSize(11);
      doc.text(`Original amount: € ${row["Original amount"]}`, 20, y); y += 10;
      doc.text(`Final amount: € ${row["Final amount"]}`, 20, y); y += 10;
      doc.text(`Commission: € ${row["Commission amount"]}`, 20, y); y += 10;

      // Booking Summary
      y += 5;
      doc.setFontSize(12);
      doc.text('--- Booking Summary ---', 20, y); y += 10;
      doc.setFontSize(11);
      doc.text('Twin Beds Room - Pool View', 20, y); y += 10;
      doc.text('Breakfast included', 20, y); y += 10;

      const detailText = 'Non-refundable (STANDARD RATE), United States (Non-refundable (STANDARD RATE) -10%)';
      doc.text(detailText, 20, y); y += 10;
      doc.text(`1 x € ${row["Final amount"]}`, 20, y); y += 10;

      doc.setFontSize(12);
      doc.text(`Total unit/room price: € ${row["Final amount"]}`, 20, y); y += 10;

      doc.setFontSize(10);
      doc.text('Rate includes 20% VAT', 20, y); y += 10;

      // Additional info
      y += 5;
      doc.setFontSize(11);
      doc.text(`Currency: ${row["Currency"]}`, 20, y);
      doc.text(`City: ${row["City"]}`, 120, y); y += 10;
      doc.text(`Property: ${row["Property name"]}`, 20, y); y += 10;

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
