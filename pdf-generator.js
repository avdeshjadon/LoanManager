/*
Copyright (c) 2025 Avdesh Jadon (LoanManager)
All Rights Reserved.
Proprietary and Confidential – Unauthorized copying, modification, or distribution of this file,
via any medium, is strictly prohibited without prior written consent from Avdesh Jadon.
*/

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

  const formatCurrencyPDF = (amount) => {
    const value = Number(amount || 0);
    return `Rs. ${value.toLocaleString("en-IN", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    })}`;
  };

  const totalPaid = paymentSchedule.reduce((sum, p) => sum + p.amountPaid, 0);

  // Use the correct interest calculation function
  const totalInterestPayable = calculateTotalInterest(
    loanDetails.principal,
    loanDetails.interestRate,
    loanDetails.loanGivenDate,
    loanDetails.loanEndDate
  );
  const totalRepayable = loanDetails.principal + totalInterestPayable;
  const outstanding = totalRepayable - totalPaid;

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
  const totalLoanBgColor = "#e6e9f8"; // Light background for the new totals section

  const paidBgColor = "#dcfce7";
  const paidTextColor = "#166534";
  const pendingBgColor = "#fef3c7";
  const pendingTextColor = "#92400e";
  const dueTextColor = "#b91c1c";

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
  doc.text("Global Finance Consultant", pageWidth / 2, y, {
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
  const contentRows = 7; 
  // Determine box height based on maximum required content, including the new KYC fields
  const kycRows = 5;
  const leftContentHeight = headerHeight + contentRows * rowHeight + 15;
  const rightContentHeight = headerHeight + (3 * rowHeight) + 15 + headerHeight + kycRows * rowHeight; // Account Summary (3 rows) + KYC Header + 5 KYC rows
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
    doc.text(value || "N/A", startX + boxWidth - 15, startY, { align: "right" });
    return startY + rowHeight;
  };

  // --- Left Box: Customer & Loan Details ---
  leftY += 25;
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(brandColor);
  doc.text("Customer & Loan Details", leftBoxX + 15, leftY);
  leftY += 15;
  
  leftY = drawDetailRow("Customer Name:", name, leftBoxX, leftY);
  leftY = drawDetailRow(
    "Loan Given Date:",
    loanDetails.loanGivenDate || "N/A",
    leftBoxX,
    leftY
  );
  leftY = drawDetailRow(
    "First Collection:",
    loanDetails.firstCollectionDate,
    leftBoxX,
    leftY
  );
  leftY = drawDetailRow(
    "Principal Amount:",
    formatCurrencyPDF(loanDetails.principal),
    leftBoxX,
    leftY
  );
  leftY = drawDetailRow(
    "Interest Rate:",
    `${loanDetails.interestRate}% Monthly`,
    leftBoxX,
    leftY
  );
  leftY = drawDetailRow(
    "Tenure:",
    `${paymentSchedule.length} ${loanDetails.frequency} installments`,
    leftBoxX,
    leftY
  );
  leftY = drawDetailRow(
    "Total Repayable:",
    formatCurrencyPDF(totalRepayable),
    leftBoxX,
    leftY
  );


  // --- Right Box: Account Summary (Current Loan) ---
  rightY += 25;
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(brandColor);
  doc.text("Account Summary (Current Loan)", rightBoxX + 15, rightY);
  rightY += 15;
  rightY = drawDetailRow(
    "Total Amount Paid:",
    formatCurrencyPDF(totalPaid),
    rightBoxX,
    rightY
  );
  rightY = drawDetailRow(
    "Outstanding Balance:",
    formatCurrencyPDF(outstanding),
    rightBoxX,
    rightY
  );
  rightY = drawDetailRow(
    "Total Interest Payable:",
    formatCurrencyPDF(totalInterestPayable),
    rightBoxX,
    rightY
  );
  
  // --- NEW: KYC & Bank Details in Right Box (Aligned Below Summary) ---
  rightY += 15; // Extra space after summary
  
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(brandColor);
  doc.text("KYC & Bank Details", rightBoxX + 15, rightY);
  rightY += 15;

  rightY = drawDetailRow("Aadhar No:", customer.aadharNumber || "N/A", rightBoxX, rightY);
  rightY = drawDetailRow("PAN No:", customer.panNumber || "N/A", rightBoxX, rightY);
  rightY = drawDetailRow("Bank Name:", customer.bankName || "N/A", rightBoxX, rightY);
  rightY = drawDetailRow("IFSC Code:", customer.ifsc || "N/A", rightBoxX, rightY);
  rightY = drawDetailRow("Account No:", customer.accountNumber || "N/A", rightBoxX, rightY);


  // --- NEW SECTION: Customer Loan Totals (All Loans) ---
  y = boxStartY + boxHeight + 20; // Start 20pt below the main boxes
  
  // Logic to calculate All Loan Totals (This relies on window.allCustomers, which must be loaded)
  const allLoans = [...window.allCustomers.active, ...window.allCustomers.settled]
      .filter(c => c.name === name && c.loanDetails && c.paymentSchedule);
  let totalP = 0, totalI = 0, totalPaidAll = 0;
  allLoans.forEach(c => {
      const li = c.loanDetails;
      // Recalculate interest for each loan
      const interest = calculateTotalInterest(li.principal, li.interestRate, li.loanGivenDate, li.loanEndDate);
      totalP += Number(li.principal || 0);
      totalI += Number(interest || 0);
      totalPaidAll += c.paymentSchedule.reduce((s, p) => s + (p.amountPaid || 0), 0);
  });
  const totalOutstandingAll = Math.max(0, totalP + totalI - totalPaidAll);

  const totalBoxX = 40;
  const totalBoxWidth = pageWidth - 80;
  const totalRows = 3;
  const totalBoxHeight = headerHeight + totalRows * rowHeight + 15;
  const totalBoxStartY = y;

  // Draw the full-width box for all loans totals
  doc.setFillColor(totalLoanBgColor);
  doc.setDrawColor(brandColor);
  doc.setLineWidth(1);
  doc.rect(totalBoxX, totalBoxStartY, totalBoxWidth, totalBoxHeight, "FD");
  
  doc.setFillColor(brandColor);
  doc.rect(totalBoxX, totalBoxStartY, totalBoxWidth, 5, "F");

  let totalY = totalBoxStartY;
  totalY += 25;
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(brandColor);
  doc.text("Customer Loan Totals (All Loans)", totalBoxX + 15, totalY);
  totalY += 15;
  
  // Use a modified drawDetailRow for the single wide box
  const drawWideDetailRow = (label, value, startY) => {
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.setTextColor(mutedColor);
    doc.text(label, totalBoxX + 15, startY);

    doc.setFont("helvetica", "bold");
    doc.setFontSize(9);
    doc.setTextColor(headingColor);
    doc.text(value || "N/A", totalBoxX + totalBoxWidth - 15, startY, { align: "right" });
    return startY + rowHeight;
  };
  
  totalY = drawWideDetailRow("Total Principal (All Loans):", formatCurrencyPDF(totalP), totalY);
  totalY = drawWideDetailRow("Total Interest (All Loans):", formatCurrencyPDF(totalI), totalY);
  totalY = drawWideDetailRow("Outstanding (All Loans):", formatCurrencyPDF(totalOutstandingAll), totalY);

  // Set Y position for the start of the next section (table)
  y = totalBoxStartY + totalBoxHeight + 30;

  // --- Table Title ---
  doc.setFontSize(14);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(headingColor);
  doc.text("Installment Repayment Schedule", 40, y);
  y += 25;

  const tableHead = [
    ["#", "Due Date", "Amount Due", "Amount Paid", "Pending", "Status"],
  ];
  const tableBody = paymentSchedule.map((inst) => [
    inst.installment,
    inst.dueDate,
    formatCurrencyPDF(inst.amountDue),
    formatCurrencyPDF(inst.amountPaid),
    formatCurrencyPDF(inst.pendingAmount),
    inst.status,
  ]);

  // --- Table Generation ---
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
      lineColor: brandColor,
    },
    styles: {
      font: "helvetica",
      fontSize: 9,
      cellPadding: { top: 8, right: 5, bottom: 8, left: 5 },
      lineColor: borderColor,
      lineWidth: 0.5,
      valign: "middle",
    },
    didParseCell: function (data) {
      if (data.row.section === "body") {
        const status = data.row.raw[5];

        if (status === "Paid") {
          data.cell.styles.fillColor = paidBgColor;
        } else if (status === "Pending") {
          data.cell.styles.fillColor = pendingBgColor;
        }

        if (data.column.dataKey === 5) {
          data.cell.styles.fontStyle = "bold";
          if (status === "Paid") {
            data.cell.styles.textColor = paidTextColor;
          } else if (status === "Pending") {
            data.cell.styles.textColor = pendingTextColor;
          } else if (status === "Due") {
            data.cell.styles.textColor = dueTextColor;
          }
        }
      }
    },
    margin: { left: 40, right: 40 } // Added margin to ensure table fits
  });

  // --- Footer ---
  const pageCount = doc.internal.getNumberOfPages();
  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i);
    doc.setFontSize(8);
    doc.setTextColor(mutedColor);
    const footerText =
      "This is a computer-generated statement and does not require a signature.";
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