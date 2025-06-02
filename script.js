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

    const logoDataUrl = await fetch('logo.png')
      .then(res => res.blob())
      .then(blob => new Promise(resolve => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.readAsDataURL(blob);
      }));

    const okReservations = rows.filter(r => r.Status === 'OK');

    for (const row of okReservations) {
      const doc = new jsPDF();
      doc.setFont('helvetica');

      // Logo
      doc.addImage(logoDataUrl, 'PNG', 10, 10, 50, 15);

      // Booking number
      doc.setFontSize(10);
      doc.text('Booking number:', 150, 15);
      doc.text(`${row["Reservation number"]}`, 150, 20);

      // Separator
      doc.setDrawColor(150);
      doc.line(10, 30, 200, 30);

      // Guest info left
      doc.setFontSize(11);
      let y = 40;
      doc.text('Guest information:', 10, y); y += 6;
      doc.setFont('helvetica', 'bold');
      doc.text(`${row["Booker name"] || ''}`, 10, y); y += 6;
      doc.setFont('helvetica', 'normal');
      doc.text(`${row["Country"] || ''}`, 10, y); y += 6;
      doc.text('Total guests:', 10, y); doc.text(`${row["Persons"]}`, 50, y); y += 6;
      doc.text('Total units/rooms:', 10, y); doc.text(`${row["Rooms"]}`, 50, y); y += 6;
      doc.text('Preferred language:', 10, y); doc.text('English', 50, y); y += 6;
      doc.text('Approx. arrival time:', 10, y); doc.text('No time provided', 50, y);

      // Check-in/out right side
      y = 40;
      doc.text('Check-in:', 130, y); doc.text(`${formatDate(row["Arrival"])}`, 170, y); y += 6;
      doc.text('Check-out:', 130, y); doc.text(`${formatDate(row["Departure"])}`, 170, y); y += 6;
      const nights = row["Room nights"] || 1;
      doc.text('Length of stay:', 130, y); doc.text(`${nights} night${nights > 1 ? 's' : ''}`, 170, y);

      // Line separator
      doc.line(10, 85, 200, 85);

      // Pricing left
      y = 95;
      doc.text('Total price:', 10, y);
      doc.setFont('helvetica', 'bold');
      doc.text(`€ ${row["Final amount"]}`, 50, y);
      doc.setFont('helvetica', 'normal');

      // Commission right
      doc.text('Commission:', 150, y); doc.text(`€ ${row["Commission amount"]}`, 180, y); y += 6;
      doc.text('Commissionable amount:', 130, y); doc.text(`€ ${row["Original amount"]}`, 180, y);

      // Separator
      doc.line(10, 110, 200, 110);

      // Room details
      y = 120;
      doc.setFont('helvetica', 'bold');
      doc.text('Room Type:', 10, y);
      doc.setFont('helvetica', 'normal');
      doc.text('Breakfast included', 10, y + 6);

      // Line for rate info
      y += 16;
      doc.text('02 - 03 Jun 2025   Non refundable (STANDARD RATE), United States (Non refundable -10%)', 10, y);
      doc.text(`1 x € ${row["Original amount"]}`, 170, y);

      y += 10;
      doc.setFont('helvetica', 'bold');
      doc.text(`Total unit/room price € ${row["Final amount"]}`, 10, y); y += 6;
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      doc.text('Rate includes 20% VAT', 10, y);

      // Save to ZIP
      const pdfBlob = doc.output('blob');
      const buffer = await pdfBlob.arrayBuffer();
      const safeName = `${(row["Booker name"] || "invoice").replace(/[^a-z0-9]/gi, '_')}_${row["Reservation number"] || Date.now()}.pdf`;
      JSZipInstance.file(safeName, buffer);
    }

    const zipBlob = await JSZipInstance.generateAsync({ type: 'blob' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(zipBlob);
    link.download = 'client_invoices.zip';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  reader.readAsArrayBuffer(file);
}

// Format date to: Mon. 3 Jun 2025
function formatDate(dateStr) {
  if (!dateStr) return '';
  const date = new Date(dateStr);
  return date.toLocaleDateString('en-GB', {
    weekday: 'short',
    day: 'numeric',
    month: 'short',
    year: 'numeric'
  });
}
