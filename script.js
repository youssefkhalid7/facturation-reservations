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

      doc.addImage(logoDataUrl, 'PNG', 10, 10, 40, 20);
      doc.setFontSize(18);
      doc.text('FACTURE CLIENT', 70, 20);

      doc.setFontSize(11);
      let y = 40;

      const fields = [
        ['Reservation number', 'Numéro de réservation'],
        ['Invoice number', 'Numéro de facture'],
        ['Booked on', 'Réservé le'],
        ['Arrival', 'Arrivée'],
        ['Departure', 'Départ'],
        ['Booker name', 'Nom du réservataire'],
        ['Guest name', 'Nom du client'],
        ['Rooms', 'Nombre de chambres'],
        ['Persons', 'Nombre de personnes'],
        ['Room nights', 'Nuits'],
        ['Commission %', 'Pourcentage de commission'],
        ['Original amount', 'Montant original'],
        ['Final amount', 'Montant final'],
        ['Commission amount', 'Montant de la commission'],
        ['Guest request', 'Demande spéciale'],
        ['Currency', 'Devise'],
        ['Hotel id', 'ID de l\'hôtel'],
        ['Property name', 'Nom de l\'établissement'],
        ['City', 'Ville'],
        ['Country', 'Pays']
      ];

      for (const [col, label] of fields) {
        doc.text(`${label}: ${row[col]}`, 20, y);
        y += 10;
      }

      // Ajout du récapitulatif
      y += 5;
      doc.setFontSize(12);
      doc.text('--- Récapitulatif de la réservation ---', 20, y); y += 10;
      doc.setFontSize(11);
      doc.text('Chambre Lits Jumeaux - Vue sur Piscine', 20, y); y += 10;
      doc.text('Petit-déjeuner inclus', 20, y); y += 10;

      const detailText = 'Non remboursable (TARIF STANDARD), United States (Non remboursable (TARIF STANDARD) -10%)';
      doc.text(detailText, 20, y); y += 10;
      doc.text(`1 x € ${row["Final amount"]}`, 20, y); y += 10;

      doc.setFontSize(12);
      doc.text(`Prix total par unité/chambre: € ${row["Final amount"]}`, 20, y); y += 10;
      doc.setFontSize(10);
      doc.text('Inclut TVA 20%', 20, y); y += 10;

      // Taxe de séjour (si applicable)
      const taxeSejour = 2.5 * (parseInt(row['Persons'] || 1));
      doc.text(`Taxe de séjour: € ${taxeSejour.toFixed(2)}`, 20, y); y += 10;

      // Enregistrement dans le ZIP
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
