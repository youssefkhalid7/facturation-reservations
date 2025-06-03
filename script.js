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
      let y = 40;

      // Logo
      doc.addImage(logoDataUrl, 'PNG', 40, y, 150, 50);
      y += 70;

      doc.setFontSize(10);
      doc.setTextColor(100);
      doc.text('Numéro de réservation :', 40, y);
      doc.setTextColor(0);
      doc.setFont(undefined, 'bold');
      doc.text(row["Reservation number"].toString(), 180, y);
      doc.setFont(undefined, 'normal');
      y += 25;

      doc.setDrawColor(200);
      doc.setLineWidth(0.5);
      doc.line(40, y, 550, y);
      y += 15;

      doc.setFontSize(12);
      doc.setTextColor(100);
      doc.text('Informations client :', 40, y);
      y += 18;
      doc.setFontSize(11);
      doc.setTextColor(0);
      doc.setFont(undefined, 'bold');
      doc.text(row["Guest name"], 40, y);
      y += 15;
      doc.setFont(undefined, 'normal');
      doc.text(row["Country"], 40, y);
      y += 15;
      doc.text(`Nombre total de personnes : ${row["Persons"]}`, 40, y);
      y += 15;
      doc.text(`Nombre total d'unités/chambres : ${row["Rooms"]}`, 40, y);
      y += 15;
      doc.text(`Langue préférée : Anglais`, 40, y);
      y += 15;
      doc.text(`Heure d’arrivée approximative : Aucune heure fournie`, 40, y);
      y += 25;

      doc.setDrawColor(200);
      doc.line(40, y, 550, y);
      y += 15;

      doc.setFontSize(10);
      doc.setTextColor(100);
      doc.text('Arrivée :', 40, y);
      doc.setTextColor(0);
      doc.setFont(undefined, 'bold');
      doc.text(row["Arrival"], 100, y);
      y += 15;
      doc.setFont(undefined, 'normal');
      doc.setTextColor(100);
      doc.text('Départ :', 40, y);
      doc.setTextColor(0);
      doc.setFont(undefined, 'bold');
      doc.text(row["Departure"], 100, y);
      y += 15;
      doc.setFont(undefined, 'normal');
      doc.setTextColor(100);
      doc.text('Durée du séjour :', 40, y);
      doc.setTextColor(0);
      doc.text(`${row["Room nights"]} nuit`, 130, y);
      y += 25;

      doc.setFontSize(12);
      doc.setTextColor(0);
      doc.setFont(undefined, 'normal');
      doc.text('Prix total :', 40, y);
      doc.setFont(undefined, 'bold');
      doc.text(`€ ${row["Final amount"]}`, 130, y);
      y += 25;

      doc.setFont(undefined, 'normal');
      doc.text('Commission :', 40, y);
      doc.setFont(undefined, 'bold');
      doc.text(`€ ${row["Commission amount"]}`, 130, y);
      y += 15;
      doc.setFont(undefined, 'normal');
      doc.text('Montant commissionnable :', 40, y);
      doc.setFont(undefined, 'bold');
      doc.text(`€ ${row["Original amount"]}`, 200, y);
      y += 20;

      doc.setFontSize(12);
      doc.setFont(undefined, 'bold');
      doc.text('Chambre Lits Jumeaux - Vue sur Piscine', 40, y);
      y += 15;
      doc.text('Petit déjeuner inclus', 40, y);
      y += 15;

      const detailText = '02 - 03 Jun 2025 Non remboursable (TARIF STANDARD), United States (Non remboursable (TARIF STANDARD) -10%) 1 x € 68,85';
      doc.setFontSize(11);
      doc.setFont(undefined, 'normal');
      doc.text(detailText, 40, y, { maxWidth: 500 });
      y += 30;

      doc.setFontSize(11);
      doc.setFont(undefined, 'normal');
      doc.text('Taxe de séjour', 40, y);
      y += 15;
      doc.text('€ 2,50   par personne et par nuit', 40, y);
      y += 15;
      const taxeTotale = (2.5 * parseInt(row["Persons"]) || 0).toFixed(2);
      doc.text(`€ ${taxeTotale}`, 40, y);
      y += 20;

      doc.text(`Prix total de l’unité/chambre € ${row["Final amount"]}`, 40, y);
      y += 15;
      doc.text('Rate includes 20 % de TVA', 40, y);

      y += 20;
      doc.setDrawColor(200);
      doc.line(40, y, 550, y);

      const pdfBlob = doc.output('blob');
      const buffer = await pdfBlob.arrayBuffer();
      const safeName = `${(row["Booker name"] || "invoice").replace(/[^a-z0-9]/gi, '_')}_${row["Reservation number"] || Date.now()}.pdf`;
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
