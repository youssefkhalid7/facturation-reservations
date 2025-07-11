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

  // 🔁 Étape 1 : Obtenir le taux EUR → MAD
  let eurToMadRate = 11; // Taux par défaut si l'API échoue
  let apiSuccess = false;

  try {
    const res = await fetch('https://v6.exchangerate-api.com/v6/26045abf677c53511c1610ce/latest/EUR');
    const data = await res.json();
    if (data.result === "success" && data.conversion_rates && data.conversion_rates.MAD) {
      eurToMadRate = data.conversion_rates.MAD;
      apiSuccess = true;
    }
  } catch (error) {
    console.warn("Erreur API. Taux par défaut utilisé.");
  }

    for (const row of rows.filter(r => r.Status === 'OK')) {
      const templateBytes = await fetch('facture_template-new.pdf').then(res => res.arrayBuffer());
      const pdfDoc = await PDFLib.PDFDocument.load(templateBytes);
      const helveticaFont = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);
      const helveticaBold = await pdfDoc.embedFont(PDFLib.StandardFonts.HelveticaBold);
      const page = pdfDoc.getPages()[0];
      const pageWidth = page.getWidth();

      // Fonctions utilitaires
      const rightAlign = (text, xRight, y, size = 12, font = helveticaFont) => {
        const textWidth = font.widthOfTextAtSize(String(text), size);
        page.drawText(String(text), {
          x: xRight - textWidth,
          y,
          size,
          font,
        });
      };
      const draw = (text, x, y, size = 12, font = helveticaFont) => {
        page.drawText(String(text), {
          x,
          y,
          size,
          font,
        });
      };

      // Booking number (gauche)
      draw(row["Reservation number"], 40, 625, 12);

      // Check-in (droite)
      rightAlign(row["Arrival"], pageWidth - 40, 625, 12);

      // Check-out (droite)
      rightAlign(row["Departure"], pageWidth - 40, 585, 12);

      // Guest information (gauche)
      const guestName = (row["Guest name"] || "").split(/[\r\n]+/)[0];
      draw(guestName, 40, 580, 12);

      // Total guests (gauche)
      draw(row["Persons"], 40, 530, 12);

      // Total units/rooms (gauche)
      draw(row["Rooms"], 40, 480, 12);

      // Length of stay (droite)
      rightAlign(`${row["Room nights"]} night${parseInt(row["Room nights"]) > 1 ? "s" : ""}`, pageWidth - 40, 530, 12);

    // 🔁 Conversion EUR → MAD si nécessaire
    let amount = parseFloat(row["Final amount"]);
    let currency = row["Currency"]?.trim().toUpperCase() || "MAD";
    let displayAmount = "";

    // Si la devise est EUR et l'API a réussi
    if (currency === "EUR" && apiSuccess) {
      amount *= eurToMadRate;
      currency = "MAD";
      displayAmount = `DH ${amount.toFixed(2)}`;
    } else if (currency === "EUR") {
      // Si la devise est EUR et l'API a échoué, on laisse EUR
      displayAmount = `€ ${amount.toFixed(2)}`;
    } else {
      // Si la devise est déjà MAD
      displayAmount = `DH ${amount.toFixed(2)}`;
    }

    // Affichage montant final (bas de la facture)
    rightAlign(displayAmount, pageWidth - 40, 340, 14, helveticaBold);

      // Total unit/room price (en bas, gras)
      // rightAlign(`€ ${parseFloat(row["Final amount"]).toFixed(2)}`, pageWidth - 40, 340, 14, helveticaBold);

      // --- Génération du PDF ---
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
