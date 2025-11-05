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

  // Find the primary loan details to populate the top summary box.
  const primaryLoan = [
    ...window.allCustomers.active,
    ...window.allCustomers.settled,
  ].find((c) => c.id === customerId);

  if (
    !primaryLoan ||
    !primaryLoan.loanDetails ||
    !Array.isArray(primaryLoan.paymentSchedule) ||
    primaryLoan.paymentSchedule.length === 0
  ) {
    alert(
      "Cannot generate PDF. The customer has incomplete or corrupt primary loan data."
    );
    return;
  }
  
  // Destructure from primaryLoan for the initial summary section
  const { loanDetails, paymentSchedule, name } = primaryLoan;


  const formatCurrencyPDF = (amount) => {
    const value = Number(amount || 0);
    return `Rs. ${value.toLocaleString("en-IN", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    })}`;
  };
  
  // Assuming 'calculateTotalInterest' is defined elsewhere in your application environment.
  if (typeof calculateTotalInterest === 'undefined') {
      alert("Error: 'calculateTotalInterest' function is missing. Cannot calculate interest totals.");
      return;
  }

  // --- Primary Loan Calculations (for Top Summary Box) ---
  const totalPaid = paymentSchedule.reduce((sum, p) => sum + (p.amountPaid || 0), 0);
  const totalInterestPayable = calculateTotalInterest(
    loanDetails.principal,
    loanDetails.interestRate,
    loanDetails.loanGivenDate,
    loanDetails.loanEndDate
  );
  const totalRepayable = Number(loanDetails.principal || 0) + Number(totalInterestPayable || 0);
  const outstanding = Math.max(0, totalRepayable - totalPaid);

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
  const totalLoanBgColor = "#e6e9f8"; // Light background for the totals and individual loan box

  const paidBgColor = "#dcfce7";
  const paidTextColor = "#166534";
  const pendingBgColor = "#fef3c7";
  const pendingTextColor = "#92400e";
  const dueTextColor = "#b91c1c"; // Used for Outstanding/Due highlights

  doc.setFont("helvetica", "normal");

  const pageHeight = doc.internal.pageSize.height;
  const pageWidth = doc.internal.pageSize.width;
  let y = 40; // Main vertical tracker

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

  // --- SECTION 1: Summary Boxes (Primary Loan & KYC) ---
  y += 30;
  const boxStartY = y;
  const leftBoxX = 40;
  const boxWidth = (pageWidth - 80 - 20) / 2;
  const rightBoxX = leftBoxX + boxWidth + 20;

  const rowHeight = 18;
  const headerHeight = 40;
  const contentRows = 7; 
  const kycRows = 5;
  const leftContentHeight = headerHeight + contentRows * rowHeight + 15;
  const rightContentHeight = headerHeight + (3 * rowHeight) + 15 + headerHeight + kycRows * rowHeight; 
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

  // --- Left Box: Customer & Primary Loan Details ---
  leftY += 25;
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(brandColor);
  doc.text("Customer & Primary Loan Details", leftBoxX + 15, leftY);
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


  // --- Right Box: Account Summary (Primary Loan) ---
  rightY += 25;
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(brandColor);
  doc.text("Account Summary (Primary Loan)", rightBoxX + 15, rightY);
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
  
  // --- KYC & Bank Details in Right Box ---
  rightY += 15; 
  
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


  // --- SECTION 2: Customer Loan Totals (All Loans) ---
  y = boxStartY + boxHeight + 20; // Start 20pt below the main boxes
  
  // Logic to calculate All Loan Totals 
  const allLoans = [...window.allCustomers.active, ...window.allCustomers.settled]
      .filter(c => c.name === name && c.loanDetails && c.paymentSchedule && c.paymentSchedule.length > 0);
      
  let totalP = 0, totalI = 0, totalPaidAll = 0;
  allLoans.forEach(c => {
      const li = c.loanDetails;
      // Recalculate interest for each loan
      const interest = calculateTotalInterest(li.principal, li.interestRate, li.loanGivenDate, li.loanEndDate);
      totalP += Number(li.principal || 0);
      totalI += Number(interest || 0);
      totalPaidAll += c.paymentSchedule.reduce((s, p) => s + (p.amountPaid || 0), 0);
  });
  const totalRepayableAll = totalP + totalI;
  const totalOutstandingAll = Math.max(0, totalRepayableAll - totalPaidAll);

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
  
  // Draw wide detail rows for All Loans summary
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
  totalY = drawWideDetailRow("Total Repayable (All Loans):", formatCurrencyPDF(totalRepayableAll), totalY);
  totalY = drawWideDetailRow("Outstanding Balance (All Loans):", formatCurrencyPDF(totalOutstandingAll), totalY);

  // Set Y position for the start of the next section (individual loan tables)
  let currentY = totalBoxStartY + totalBoxHeight + 30;

  
  // --- SECTION 3: Individual Loan Repayment Schedules (With New Box Layout) ---
  
  doc.setFontSize(14);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(headingColor);
  doc.text("Individual Loan Repayment Schedules", 40, currentY);
  currentY += 25;

  // Loop through each valid loan the customer has
  allLoans.forEach((loan, index) => {
    
    const { loanDetails: currentLoanDetails, paymentSchedule: currentPaymentSchedule } = loan;

    // --- Page Break Check ---
    // If the space left on the page is less than 180pt, add a new page (180pt accounts for the box + table start)
    if (currentY + 180 > pageHeight - 50 && index > 0) { 
        doc.addPage();
        currentY = 40; // Reset Y for the new page
    }

    // Recalculate Totals for this specific loan for the header
    const currentTotalInterestPayable = calculateTotalInterest(
        currentLoanDetails.principal,
        currentLoanDetails.interestRate,
        currentLoanDetails.loanGivenDate,
        currentLoanDetails.loanEndDate
    );
    const currentTotalRepayable = Number(currentLoanDetails.principal || 0) + Number(currentTotalInterestPayable || 0);
    const currentTotalPaid = currentPaymentSchedule.reduce((sum, p) => sum + (p.amountPaid || 0), 0);
    const currentOutstanding = Math.max(0, currentTotalRepayable - currentTotalPaid);

    // --- Loan Title (ID Removed) ---
    doc.setFontSize(12);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(brandColor);
    doc.text(`Loan ${index + 1} Statement`, 40, currentY);
    currentY += 20;

    // --- NEW: Individual Loan Summary Box (6 fields in 3 rows) ---
    const loanBoxX = 40;
    const loanBoxWidth = pageWidth - 80;
    const loanBoxHeight = 63; // Height for three rows of details (3 * 18pt row height + padding)
    const loanBoxStartY = currentY;
    const padding = 10;
    
    // Draw the box background
    doc.setFillColor(totalLoanBgColor);
    doc.setDrawColor(borderColor);
    doc.setLineWidth(1);
    doc.rect(loanBoxX, loanBoxStartY, loanBoxWidth, loanBoxHeight, "FD");
    
    let boxInnerY = loanBoxStartY + 18;
    
    // Helper function to draw a label/value pair in the box (two pairs per row)
    const drawLoanDetailBoxItem = (label, value, startX, startY, isOutstanding = false) => {
        const colWidth = loanBoxWidth / 2;
        
        doc.setFont("helvetica", "normal");
        doc.setFontSize(8.5);
        doc.setTextColor(mutedColor);
        
        // Label position
        doc.text(label, startX + padding, startY);

        // Value position
        doc.setFont(isOutstanding ? "helvetica" : "helvetica", isOutstanding ? "bold" : "normal");
        doc.setFontSize(9);
        doc.setTextColor(isOutstanding ? dueTextColor : headingColor); 
        
        // Value aligns right within its half column
        doc.text(value || "N/A", startX + colWidth - padding, startY, { align: "right" });
    };
    
    const col1X = loanBoxX;
    const col2X = loanBoxX + loanBoxWidth / 2;

    // Row 1: Principal and Repayable
    drawLoanDetailBoxItem("Principal Amount:", formatCurrencyPDF(currentLoanDetails.principal), col1X, boxInnerY);
    drawLoanDetailBoxItem("Total Repayable:", formatCurrencyPDF(currentTotalRepayable), col2X, boxInnerY);
    boxInnerY += 18; 

    // Row 2: Rate and Outstanding
    drawLoanDetailBoxItem("Interest Rate:", `${currentLoanDetails.interestRate}% Monthly`, col1X, boxInnerY);
    drawLoanDetailBoxItem("Outstanding Balance:", formatCurrencyPDF(currentOutstanding), col2X, boxInnerY, true); 
    boxInnerY += 18;

    // Row 3: Dates (New Addition)
    drawLoanDetailBoxItem("Loan Given Date:", currentLoanDetails.loanGivenDate || "N/A", col1X, boxInnerY);
    drawLoanDetailBoxItem("Loan End Date:", currentLoanDetails.loanEndDate || "N/A", col2X, boxInnerY);
    

    // Update currentY for the table
    currentY = loanBoxStartY + loanBoxHeight + 15; // 15pt space after the box

    // --- Table Generation for THIS Loan ---
    
    const tableHead = [
        ["#", "Due Date", "Amount Due", "Amount Paid", "Pending", "Status"],
    ];
    const tableBody = currentPaymentSchedule.map((inst) => [
        inst.installment,
        inst.dueDate,
        formatCurrencyPDF(inst.amountDue),
        formatCurrencyPDF(inst.amountPaid),
        formatCurrencyPDF(inst.pendingAmount),
        inst.status,
    ]);

    // Use autoTable to place the table
    doc.autoTable({
        head: tableHead,
        body: tableBody,
        startY: currentY, // Use currentY as the starting point
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
        margin: { left: 40, right: 40 },
        
        // Update currentY position after the table is drawn/page is completed
        didDrawPage: function (data) {
             currentY = data.cursor.y + 30; 
        },
    });

    // Update currentY using the final position of the last element drawn by autoTable
    if (doc.autoTable.previous.finalY) {
        currentY = doc.autoTable.previous.finalY + 30; // Add space before the next loan/footer
    }
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