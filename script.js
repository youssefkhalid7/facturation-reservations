async function processExcel() {
    const fileInput = document.getElementById('excelFile');
    if (!fileInput.files[0]) return alert('Veuillez téléverser un fichier Excel.');
  
    const file = fileInput.files[0];
    const reader = new FileReader();
  
    reader.onload = async function(e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet);
  
      const JSZipInstance = new JSZip();
      const { jsPDF } = window.jspdf;
  
      const okReservations = rows.filter(r => r.Status === 'OK');
  
      for (const row of okReservations) {
        const doc = new jsPDF();
        doc.text("FACTURE CLIENT", 90, 20);
        doc.setFontSize(10);
  
        const champs = [
          "Reservation number", "Invoice number", "Booked on", "Arrival", "Departure",
          "Booker name", "Guest name", "Rooms", "Persons", "Room nights",
          "Commission %", "Original amount", "Final amount", "Commission amount",
          "Status", "Guest request", "Currency", "Hotel id", "Property name", "City", "Country"
        ];
  
        let y = 40;
        for (const champ of champs) {
          const label = champ.replace(/_/g, ' ');
          const valeur = row[champ] !== undefined ? row[champ] : '';
          doc.text(`${label}: ${valeur}`, 20, y);
          y += 10;
          if (y > 270) {
            doc.addPage();
            y = 20;
          }
        }
  
        const pdfBlob = doc.output('blob');
        const buffer = await pdfBlob.arrayBuffer();
        const safeName = `${(row["Booker name"] || "facture").replace(/[^a-z0-9]/gi, '_')}_${row["Reservation number"] || Date.now()}.pdf`;
        JSZipInstance.file(safeName, buffer);
      }
  
      const zipBlob = await JSZipInstance.generateAsync({ type: 'blob' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(zipBlob);
      link.download = 'factures_clients.zip';
  
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    };
  
    reader.readAsArrayBuffer(file);
  }
  