/*
Copyright (c) 2025 Avdesh Jadon (LoanManager)
All Rights Reserved.
Proprietary and Confidential – Unauthorized copying, modification, or distribution of this file,
via any medium, is strictly prohibited without prior written consent from Avdesh Jadon.
*/

async function generateAndDownloadPDF(customerId) {
  // --- Date Formatting Helper ---
  const formatDate = (dateInput) => {
    if (!dateInput && dateInput !== 0) return "N/A";
    try {
      // If already a Date instance
      if (dateInput instanceof Date) {
        if (isNaN(dateInput.getTime())) return "N/A";
        const d = dateInput;
        const day = String(d.getDate()).padStart(2, "0");
        const month = String(d.getMonth() + 1).padStart(2, "0");
        const year = d.getFullYear();
        return `${day}-${month}-${year}`;
      }

      const s = String(dateInput).trim();

      // YYYY-MM-DD or YYYY/MM/DD
      if (/^\d{4}[-/]\d{2}[-/]\d{2}$/.test(s)) {
        const [y, mo, d] = s.split(/[-/]/);
        const dObj = new Date(Number(y), Number(mo) - 1, Number(d));
        if (!isNaN(dObj.getTime())) {
          return `${String(dObj.getDate()).padStart(2, "0")}-${String(dObj.getMonth() + 1).padStart(2, "0")}-${dObj.getFullYear()}`;
        }
      }

      // DD-MM-YYYY or DD/MM/YYYY
      if (/^\d{2}[-/]\d{2}[-/]\d{4}$/.test(s)) {
        const [d, mo, y] = s.split(/[-/]/);
        const dObj = new Date(Number(y), Number(mo) - 1, Number(d));
        if (!isNaN(dObj.getTime())) {
          return `${String(dObj.getDate()).padStart(2, "0")}-${String(dObj.getMonth() + 1).padStart(2, "0")}-${dObj.getFullYear()}`;
        }
      }

      // Try Date.parse for ISO / other parseable strings
      const parsed = Date.parse(s);
      if (!isNaN(parsed)) {
        const dObj = new Date(parsed);
        return `${String(dObj.getDate()).padStart(2, "0")}-${String(dObj.getMonth() + 1).padStart(2, "0")}-${dObj.getFullYear()}`;
      }

      // Fallback: try basic split heuristic and format as DD-MM-YYYY when possible
      const parts = s.split(/[-/]/);
      if (parts.length === 3) {
        if (parts[0].length === 4) {
          // YYYY-MM-DD -> convert to DD-MM-YYYY
          return `${String(Number(parts[2])).padStart(2, "0")}-${String(Number(parts[1])).padStart(2, "0")}-${parts[0]}`;
        }
        if (parts[2].length === 4) {
          // DD-MM-YYYY already -> normalize numeric padding
          return `${String(Number(parts[0])).padStart(2, "0")}-${String(Number(parts[1])).padStart(2, "0")}-${parts[2]}`;
        }
      }

      return s;
    } catch (e) {
      return String(dateInput);
    }
  };


  // 1. INITIAL CUSTOMER & PRIMARY LOAN IDENTIFICATION
  const primaryLoan = [
    ...window.allCustomers.active,
    ...window.allCustomers.settled,
  ].find((c) => c.id === customerId);

  if (!primaryLoan) {
    alert("Could not find customer data associated with this loan ID.");
    return;
  }
  
  const customer = primaryLoan; 
  const { loanDetails, paymentSchedule, name } = primaryLoan; 
  
  if (typeof calculateTotalInterest === 'undefined') {
      alert("Error: 'calculateTotalInterest' function is missing.");
      return;
  }
  
  const formatCurrencyPDF = (amount) => {
    const value = Number(amount || 0);
    return `Rs. ${value.toLocaleString("en-IN", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    })}`;
  };
  
  // --- PDF Setup ---
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: "p", unit: "pt", format: "a4" });

  // --- Color Palette ---
  const brandColor = "#4a55a2";
  const headingColor = "#2d3748";
  const mutedColor = "#718096";
  const borderColor = "#e2e8f0";
  const backgroundColor = "#f8fafc";
  const totalLoanBgColor = "#e6e9f8"; 

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
  doc.text("Global Finance Consultant", pageWidth / 2, y, { align: "center" });
  y += 20;
  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);
  doc.setTextColor(mutedColor);
  doc.text("Loan Statement", pageWidth / 2, y, {
    align: "center",
  });
  y += 15;
  doc.setLineWidth(1.5);
  doc.setDrawColor(brandColor);
  doc.line(40, y, pageWidth - 40, y);

  // --- SECTION 1: Customer Personal & KYC Details (Combined in Two Columns) ---
  y += 30;
  const boxStartY = y;
  const leftBoxX = 40;
  const boxWidth = (pageWidth - 80 - 20) / 2;
  const rightBoxX = leftBoxX + boxWidth + 20;

  const rowHeight = 18;
  const headerHeight = 40;
  
  // Adjusted rows to accommodate Address split/space
  const contentRowsLeft = 4; // Name, Aadhar, PAN, Phone
  const contentRowsRight = 7; // Email, Bank, IFSC, Account, Address Label, Address Value, Padding
  
  const boxHeight = headerHeight + contentRowsRight * rowHeight + 15; // Box height defined by the Right side's rows


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

  // Re-define drawDetailRow helper for clarity
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
  
  // NEW Helper for Address to use full width of the box and small font
  const drawAddressRow = (label, value, startX, startY) => {
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8); // Smaller font size for Address label
    doc.setTextColor(mutedColor);
    doc.text(label, startX + 15, startY);

    doc.setFont("helvetica", "normal");
    doc.setFontSize(8.5); // Slightly larger font for value
    doc.setTextColor(headingColor);
    // Align address text to the left/justified within the available space
    doc.text(value || "N/A", startX + 15, startY + 12, { maxWidth: boxWidth - 30, lineHeightFactor: 1.2 });
    return startY + rowHeight + 15; // Use extra spacing after address line
  };

  // --- Left Box: Customer Details ONLY (Name + KYC Start) ---
  leftY += 25;
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(brandColor);
  doc.text("Customer Personal Details", leftBoxX + 15, leftY);
  leftY += 15;
  
  leftY = drawDetailRow("Customer Name:", name, leftBoxX, leftY);
  leftY = drawDetailRow("Aadhar No:", customer.aadharNumber || "N/A", leftBoxX, leftY);
  leftY = drawDetailRow("PAN No:", customer.panNumber || "N/A", leftBoxX, leftY);
  leftY = drawDetailRow("Phone No:", customer.phoneNumber || "N/A", leftBoxX, leftY);

  
  // --- Right Box: Bank & Contact Details (KYC End) ---
  rightY += 25;
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(brandColor);
  doc.text("Bank & Contact Details", rightBoxX + 15, rightY);
  rightY += 15;

  rightY = drawDetailRow("Email:", customer.email || "N/A", rightBoxX, rightY);
  rightY = drawDetailRow("Bank Name:", customer.bankName || "N/A", rightBoxX, rightY);
  rightY = drawDetailRow("IFSC Code:", customer.ifsc || "N/A", rightBoxX, rightY);
  rightY = drawDetailRow("Account No:", customer.accountNumber || "N/A", rightBoxX, rightY);
  
  // --- Address moved to a dedicated section at the end of the Right Box (FIXED) ---
  rightY += 5; // Small extra space
  doc.setLineWidth(0.5);
  doc.setDrawColor(borderColor);
  doc.line(rightBoxX + 10, rightY, rightBoxX + boxWidth - 10, rightY);
  rightY += 15;
  
  // Use the new helper function for address
  rightY = drawAddressRow("Full Address:", customer.address || "N/A", rightBoxX, rightY);


  // --- SECTION 2: Customer Loan Totals (All Loans) ---
  y = boxStartY + boxHeight + 20; // Start 20pt below the main boxes
  
  // Logic to calculate All Loan Totals 
  // Build a sequence-aware sort: prefer explicit loan sequence fields, fallback to loanGivenDate
  const extractSequence = (c) => {
    const ld = c.loanDetails || {};
    const candidates = [ld.sequence, ld.loanSequence, ld.loanNumber, ld.loanNo];
    for (const v of candidates) {
      if (v != null && v !== "") {
        const n = Number(String(v).replace(/[^\d]/g, ""));
        if (!isNaN(n)) return n;
      }
    }
    return null;
  };
  
  // Robust parser that returns a Date (or null) for common formats (DD-MM-YYYY, YYYY-MM-DD, ISO)
  const parseDateForSort = (dateInput) => {
    if (!dateInput && dateInput !== 0) return null;
    if (dateInput instanceof Date) return isNaN(dateInput.getTime()) ? null : dateInput;
    const s = String(dateInput).trim();

    // DD-MM-YYYY or DD/MM/YYYY
    if (/^\d{2}[-/]\d{2}[-/]\d{4}$/.test(s)) {
      const [d, mo, y] = s.split(/[-/]/).map(Number);
      const dt = new Date(y, mo - 1, d);
      return isNaN(dt.getTime()) ? null : dt;
    }

    // YYYY-MM-DD or YYYY/MM/DD
    if (/^\d{4}[-/]\d{2}[-/]\d{2}$/.test(s)) {
      const [y, mo, d] = s.split(/[-/]/).map(Number);
      const dt = new Date(y, mo - 1, d);
      return isNaN(dt.getTime()) ? null : dt;
    }

    // Try Date.parse (ISO and other parseable strings)
    const parsed = Date.parse(s);
    if (!isNaN(parsed)) return new Date(parsed);

    return null;
  };

  const extractGivenTime = (c) => {
    const d = parseDateForSort((c.loanDetails && c.loanDetails.loanGivenDate) || null);
    return d ? d.getTime() : null;
  };

  // Build list with original index to preserve stable ordering as a final fallback
  const sourceList = [...window.allCustomers.active, ...window.allCustomers.settled];
  const filtered = sourceList.filter((c) => c.name === customer.name && c.loanDetails && c.paymentSchedule && c.paymentSchedule.length > 0);
  const withMeta = filtered.map((c, i) => ({
    item: c,
    origIndex: i,
    seq: extractSequence(c),
    time: extractGivenTime(c),
  }));
  
  withMeta.sort((a, b) => {
    if (a.seq != null && b.seq != null) return a.seq - b.seq;
    if (a.seq != null) return -1;
    if (b.seq != null) return 1;
    if (a.time != null && b.time != null) return a.time - b.time; // oldest-first
    if (a.time != null) return -1;
    if (b.time != null) return 1;
    // final fallback -> preserve original filtered order
    return a.origIndex - b.origIndex;
  });

  const allLoans = withMeta.map((m) => m.item);
      
  let totalP = 0, totalI = 0, totalPaidAll = 0;
  allLoans.forEach(c => {
      const li = c.loanDetails;
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

  
  // --- SECTION 3: Individual Loan Repayment Schedules ---
  
  doc.setFontSize(14);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(headingColor);
  doc.text("Individual Loan Repayment Schedules", 40, currentY);
  currentY += 25;

  // Loop through each valid loan the customer has (now sorted)
  allLoans.forEach((loan, index) => {
    
    const { loanDetails: currentLoanDetails, paymentSchedule: currentPaymentSchedule } = loan;

    // --- Page Break Check ---
    if (currentY + 180 > pageHeight - 50 && index > 0) { 
        doc.addPage();
        currentY = 40; 
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
    // Prefer explicit sequence display (sequence, loanSequence, loanNumber, loanNo), otherwise use index+1
    const seqValue =
      currentLoanDetails.sequence ||
      currentLoanDetails.loanSequence ||
      currentLoanDetails.loanNumber ||
      currentLoanDetails.loanNo ||
      (index + 1);
    doc.setFontSize(12);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(brandColor);
    doc.text(`Finance ${seqValue} Statement`, 40, currentY);
    currentY += 20;

    // --- Individual Loan Summary Box (5 fields, table-like format) ---
    const loanBoxX = 40;
    const loanBoxWidth = pageWidth - 80;
    const loanBoxHeight = 63; 
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

    // Row 3: Loan Given Date (Loan End Date removed)
    drawLoanDetailBoxItem("Loan Given Date:", formatDate(currentLoanDetails.loanGivenDate), col1X, boxInnerY); 
    

    // Update currentY for the table
    currentY = loanBoxStartY + loanBoxHeight + 15; 

    // --- Table Generation for THIS Loan (Due Date Formatted) ---
    
    const tableHead = [
        ["#", "Due Date", "Amount Due", "Amount Paid", "Pending", "Status"],
    ];
    const tableBody = currentPaymentSchedule.map((inst) => [
        inst.installment,
        formatDate(inst.dueDate), // Date Format Applied
        formatCurrencyPDF(inst.amountDue),
        formatCurrencyPDF(inst.amountPaid),
        formatCurrencyPDF(inst.pendingAmount),
        inst.status,
    ]);

    // Use autoTable to place the table
    doc.autoTable({
        head: tableHead,
        body: tableBody,
        startY: currentY, 
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
        
        didDrawPage: function (data) {
             currentY = data.cursor.y + 30; 
        },
    });

    if (doc.autoTable.previous.finalY) {
        currentY = doc.autoTable.previous.finalY + 30;
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