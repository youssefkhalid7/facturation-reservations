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
  
      // Charger le logo
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
  
        // Logo + Titre
        doc.addImage(logoDataUrl, 'PNG', 10, 10, 40, 20);
        doc.setFontSize(18);
        doc.text('FACTURE CLIENT', 70, 20);
  
        doc.setFontSize(11);
        let y = 40;
  
        doc.text(`Réservation: ${row["Reservation number"]}`, 20, y); y += 10;
        doc.text(`Client: ${row["Booker name"]}`, 20, y); y += 10;
        doc.text(`Pays: ${row["Country"]}`, 20, y); y += 10;
        doc.text(`Arrivée: ${row["Arrival"]}`, 20, y);
        doc.text(`Départ: ${row["Departure"]}`, 120, y); y += 10;
        doc.text(`Nombre de nuits: ${row["Room nights"]}`, 20, y); y += 10;
        doc.text(`Chambre(s): ${row["Rooms"]}`, 20, y);
        doc.text(`Personnes: ${row["Persons"]}`, 120, y); y += 10;
  
        // Détails financiers
        y += 5;
        doc.setFontSize(12);
        doc.text('--- Détails de la facturation ---', 20, y); y += 10;
        doc.setFontSize(11);
        doc.text(`Montant original: € ${row["Original amount"]}`, 20, y); y += 10;
        doc.text(`Montant final: € ${row["Final amount"]}`, 20, y); y += 10;
        doc.text(`Commission: € ${row["Commission amount"]}`, 20, y); y += 10;
  
        // Total, Taxes, Devise
        y += 5;
        doc.text(`Devise: ${row["Currency"]}`, 20, y);
        doc.text(`Ville: ${row["City"]}`, 120, y); y += 10;
        doc.text(`Propriété: ${row["Property name"]}`, 20, y); y += 10;
  
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
  