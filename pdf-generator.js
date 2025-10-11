async function generateAndDownloadPDF(customerId) {
  const customer = [
    ...window.allCustomers.active,
    ...window.allCustomers.settled,
  ].find((c) => c.id === customerId);

  if (!customer) {
    alert("Could not find customer data.");
    return;
  }

  if (
    !customer.loanDetails ||
    !Array.isArray(customer.paymentSchedule) ||
    customer.paymentSchedule.length === 0
  ) {
    alert(
      "Cannot generate PDF. The customer has incomplete or corrupt loan data."
    );
    return;
  }

  const { loanDetails, paymentSchedule, name } = customer;
  const formatCurrency = (amount) =>
    `Rs. ${Math.round(Number(amount || 0)).toLocaleString("en-IN")}`;

  const totalPaid = paymentSchedule.reduce(
    (sum, p) => sum + p.amountPaid,
    0
  );
  const totalRepayable =
    loanDetails.principal * (1 + loanDetails.interestRate / 100);
  const outstanding = totalRepayable - totalPaid;
  const totalInterest =
    loanDetails.principal * (loanDetails.interestRate / 100);

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({
    orientation: "p",
    unit: "pt",
    format: "a4",
  });

  // --- Color Palette ---
  const brandColor = "#4a55a2";
  const headingColor = "#2d3748";
  const mutedColor = "#718096";
  const borderColor = "#e2e8f0";
  const backgroundColor = "#f8fafc";
  
  const paidBgColor = "#dcfce7";      // Paid: Fresh, light green background
  const paidTextColor = "#166534";   // Paid: Dark, rich green text
  const pendingBgColor = "#fef3c7";   // Pending: Warning-type yellow background
  const pendingTextColor = "#92400e"; // Pending: Dark amber text
  const dueTextColor = "#b91c1c";     // Due: Normal background, but bold red text

  doc.setFont("helvetica", "normal");

  const pageHeight = doc.internal.pageSize.height;
  const pageWidth = doc.internal.pageSize.width;
  let y = 40;

  // --- Header ---
  doc.setFont("helvetica", "bold");
  doc.setFontSize(22);
  doc.setTextColor(headingColor);
  doc.text("Loan Statement", pageWidth / 2, y, { align: "center" });
  y += 20;
  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);
  doc.setTextColor(mutedColor);
  doc.text("Your Trusted Lending Partner", pageWidth / 2, y, {
    align: "center",
  });
  y += 15;
  doc.setLineWidth(1.5);
  doc.setDrawColor(brandColor);
  doc.line(40, y, pageWidth - 40, y);

  // --- Summary Section ---
  y += 30;
  const boxStartY = y;
  const leftBoxX = 40;
  const boxWidth = (pageWidth - 80 - 20) / 2;
  const rightBoxX = leftBoxX + boxWidth + 20;

  const rowHeight = 18;
  const headerHeight = 40;
  const leftContentHeight = headerHeight + (6 * rowHeight);
  const rightContentHeight = headerHeight + (3 * rowHeight);
  const boxHeight = Math.max(leftContentHeight, rightContentHeight) + 15;

  doc.setFillColor(backgroundColor);
  doc.setDrawColor(borderColor);
  doc.setLineWidth(1);
  doc.rect(leftBoxX, boxStartY, boxWidth, boxHeight, "FD");
  doc.rect(rightBoxX, boxStartY, boxWidth, boxHeight, "FD");

  doc.setFillColor(brandColor);
  doc.rect(leftBoxX, boxStartY, boxWidth, 5, "F");
  doc.rect(rightBoxX, boxStartY, boxWidth, 5, "F");

  let leftY = boxStartY;
  let rightY = boxStartY;
  
  const drawDetailRow = (label, value, startX, startY) => {
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.setTextColor(mutedColor);
    doc.text(label, startX + 15, startY);

    doc.setFont("helvetica", "bold");
    doc.setFontSize(9);
    doc.setTextColor(headingColor);
    doc.text(value, startX + boxWidth - 15, startY, { align: "right" });
    return startY + rowHeight;
  };

  leftY += 25;
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(brandColor);
  doc.text("Customer & Loan Details", leftBoxX + 15, leftY);
  leftY += 15;
  leftY = drawDetailRow("Customer Name:", name, leftBoxX, leftY);
  leftY = drawDetailRow("Loan Date:", loanDetails.loanDate, leftBoxX, leftY);
  leftY = drawDetailRow("Principal Amount:", formatCurrency(loanDetails.principal), leftBoxX, leftY);
  leftY = drawDetailRow("Interest Rate:", `${loanDetails.interestRate}%`, leftBoxX, leftY);
  leftY = drawDetailRow("Tenure:", `${loanDetails.installments} installments`, leftBoxX, leftY);
  leftY = drawDetailRow("Installment Amount:", formatCurrency(paymentSchedule[0].amountDue), leftBoxX, leftY);

  rightY += 25;
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(brandColor);
  doc.text("Account Summary", rightBoxX + 15, rightY);
  rightY += 15;
  rightY = drawDetailRow("Total Amount Paid:", formatCurrency(totalPaid), rightBoxX, rightY);
  rightY = drawDetailRow("Outstanding Balance:", formatCurrency(outstanding), rightBoxX, rightY);
  rightY = drawDetailRow("Total Interest:", formatCurrency(totalInterest), rightBoxX, rightY);

  y = boxStartY + boxHeight + 30;

  // --- Table Title ---
  doc.setFontSize(14);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(headingColor);
  doc.text("Installment Repayment Schedule", 40, y);
  y += 25;

  const tableHead = [["#", "Due Date", "Amount Due", "Amount Paid", "Pending", "Status"]];
  const tableBody = paymentSchedule.map((inst) => [
    inst.installment,
    inst.dueDate,
    formatCurrency(inst.amountDue),
    formatCurrency(inst.amountPaid),
    formatCurrency(inst.pendingAmount),
    inst.status,
  ]);

  // --- Table Generation with Fixed Styling Logic ---
  doc.autoTable({
    head: tableHead,
    body: tableBody,
    startY: y,
    theme: "grid",
    headStyles: {
      fillColor: brandColor,
      textColor: "#ffffff",
      font: "helvetica",
      fontStyle: "bold",
      fontSize: 9,
      halign: "center",
      lineColor: brandColor
    },
    styles: {
      font: "helvetica",
      fontSize: 9,
      cellPadding: { top: 8, right: 5, bottom: 8, left: 5 },
      lineColor: borderColor,
      lineWidth: 0.5,
      valign: 'middle'
    },
    didParseCell: function(data) {
      if (data.row.section === 'body') {
        const status = data.row.raw[5];

        // Apply background color to every cell in the row based on status
        if (status === 'Paid') {
          data.cell.styles.fillColor = paidBgColor;
        } else if (status === 'Pending') {
          data.cell.styles.fillColor = pendingBgColor;
        }

        // Apply specific text styling JUST for the status column
        if (data.column.dataKey === 5) {
          data.cell.styles.fontStyle = 'bold';
          if (status === 'Paid') {
            data.cell.styles.textColor = paidTextColor;
          } else if (status === 'Pending') {
            data.cell.styles.textColor = pendingTextColor;
          } else if (status === 'Due') {
            data.cell.styles.textColor = dueTextColor;
          }
        }
      }
    },
  });

  // --- Footer ---
  const pageCount = doc.internal.getNumberOfPages();
  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i);
    doc.setFontSize(8);
    doc.setTextColor(mutedColor);
    const footerText = "This is a computer-generated statement and does not require a signature.";
    doc.text(footerText, pageWidth / 2, pageHeight - 25, {
      align: "center",
    });
    doc.text(`Page ${i} of ${pageCount}`, pageWidth - 40, pageHeight - 25, {
      align: "right",
    });
  }

  const pdfName = `Loan-Statement-${name.replace(/ /g, "-")}.pdf`;
  doc.save(pdfName);
}