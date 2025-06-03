async function processExcel() {
  const fileInput = document.getElementById('excelFile');
  if (!fileInput.files[0]) return alert('Veuillez importer un fichier Excel.');
  const file = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = async function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    const reservations = rows.filter(r => r.Status === 'OK');

    const templateBytes = await fetch('4490222232.pdf').then(res => res.arrayBuffer());
    const zip = new JSZip();

    for (const row of reservations) {
      const pdfDoc = await PDFLib.PDFDocument.load(templateBytes);
      const form = pdfDoc.getForm();

      form.getTextField('BookingNumber').setText(row["Reservation number"].toString());
      form.getTextField('GuestName').setText(row["Guest name"]);
      form.getTextField('Country').setText(row["Country"]);
      form.getTextField('Persons').setText(`${row["Persons"]}`);
      form.getTextField('Rooms').setText(`${row["Rooms"]}`);
      form.getTextField('Arrival').setText(row["Arrival"]);
      form.getTextField('Departure').setText(row["Departure"]);
      form.getTextField('RoomNights').setText(`${row["Room nights"]}`);
      form.getTextField('FinalAmount').setText(`€ ${row["Final amount"]}`);
      form.getTextField('Commission').setText(`€ ${row["Commission amount"]}`);
      form.getTextField('OriginalAmount').setText(`€ ${row["Original amount"]}`);
      form.getTextField('Tax').setText(`€ ${(2.5 * parseInt(row["Persons"]) || 0).toFixed(2)}`);

      form.flatten(); // Verrouille les champs

      const pdfBytes = await pdfDoc.save();
      const fileName = `${(row["Booker name"] || "facture").replace(/[^a-z0-9]/gi, '_')}_${row["Reservation number"]}.pdf`;
      zip.file(fileName, pdfBytes);
    }

    const zipBlob = await zip.generateAsync({ type: 'blob' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(zipBlob);
    a.download = 'factures_clients.zip';
    a.click();
  };

  reader.readAsArrayBuffer(file);
}