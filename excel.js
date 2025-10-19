/*
Copyright (c) 2025 Avdesh Jadon (LoanManager)
All Rights Reserved.
Proprietary and Confidential – Unauthorized copying, modification, or distribution of this file,
via any medium, is strictly prohibited without prior written consent from Avdesh Jadon.
*/

function exportToExcel(customers, fileName = "customer-data.xlsx") {
  if (!customers || customers.length === 0) {
    alert("No data to export.");
    return;
  }

  const flattenedData = customers.map((customer) => {
    if (!customer.loanDetails || !Array.isArray(customer.paymentSchedule)) {
      return {
        Name: customer.name,
        Phone: customer.phone || "N/A",
        Status: "Incomplete/Corrupt Data",
      };
    }

    const totalPaid = customer.paymentSchedule.reduce(
      (sum, p) => sum + p.amountPaid,
      0
    );
    // Compute total repayable using 30-day block interest across full loan term
    let totalRepayable = customer.loanDetails.principal;
    if (typeof calculateTotalInterest === 'function') {
      try {
        const ti = calculateTotalInterest(
          customer.loanDetails.principal,
          customer.loanDetails.interestRate,
          customer.loanDetails.loanGivenDate,
          customer.loanDetails.loanEndDate
        );
        totalRepayable = customer.loanDetails.principal + ti;
      } catch (_) {
        totalRepayable = customer.loanDetails.principal * (1 + customer.loanDetails.interestRate / 100);
      }
    } else {
      totalRepayable = customer.loanDetails.principal * (1 + customer.loanDetails.interestRate / 100);
    }
    const outstanding = totalRepayable - totalPaid;
    const totalInterestPaid = Math.max(
      0,
      totalPaid - customer.loanDetails.principal
    );
    const paidInstallments = customer.paymentSchedule.filter(
      (p) => p.status === "Paid"
    ).length;

    const kycDocs = customer.kycDocs || {};

    return {
      Name: customer.name,
      "WhatsApp Number": customer.phone || "N/A",
      "Father's Name": customer.fatherName || "N/A",
      Address: customer.address || "N/A",
      "Aadhar URL": kycDocs.aadharUrl || "N/A",
      "PAN URL": kycDocs.panUrl || "N/A",
      "Picture URL": kycDocs.picUrl || "N/A",
      "Bank Details URL": kycDocs.bankDetailsUrl || "N/A",
      "Loan Principal (₹)": customer.loanDetails.principal,
      "Total Interest Rate (%)": customer.loanDetails.interestRate,
      "No. of Installments": customer.loanDetails.installments,
      Frequency: customer.loanDetails.frequency,
      "Installment Amount (₹)": customer.paymentSchedule[0]?.amountDue,
      "Loan Date": customer.loanDetails.loanDate,
      Status: customer.status,
      "Installments Paid": `${paidInstallments}/${customer.paymentSchedule.length}`,
      "Outstanding Balance (₹)": outstanding,
      "Total Amount Paid (₹)": totalPaid,
      "Total Interest Paid (₹)": totalInterestPaid,
    };
  });

  const worksheet = XLSX.utils.json_to_sheet(flattenedData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Customers");
  XLSX.writeFile(workbook, fileName);
}
