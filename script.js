const firebaseConfig = {
  apiKey: "AIzaSyD0wHBp_Gb-U7eJFjoa3pTBYdgipxMYyzg",
  authDomain: "globalfinanceconsultant-bf13a.firebaseapp.com",
  projectId: "globalfinanceconsultant-bf13a",
  storageBucket: "globalfinanceconsultant-bf13a.firebasestorage.app",
  messagingSenderId: "611257450437",
  appId: "1:611257450437:web:fda4e59be985ef146e55ac",
  measurementId: "G-KWM1STNNZ0",
};
firebase.initializeApp(firebaseConfig);
const db = firebase.firestore();
const auth = firebase.auth();
const storage = firebase.storage();

window.allCustomers = { active: [], settled: [] };

// NEW: Calculates total interest based on conceptual months, not raw days.
const calculateConceptualTotalInterest = (
  principal,
  monthlyRate,
  firstDateStr,
  endDateStr
) => {
  if (!principal || !monthlyRate || !firstDateStr || !endDateStr) return 0;

  const monthlyRateDecimal = monthlyRate / 100;

  // Calculate the number of conceptual months between the start and end dates.
  const d1 = new Date(firstDateStr);
  const d2 = new Date(endDateStr);
  let conceptualMonths = (d2.getFullYear() - d1.getFullYear()) * 12;
  conceptualMonths -= d1.getMonth();
  conceptualMonths += d2.getMonth();
  if (d2.getDate() < d1.getDate()) {
    conceptualMonths--;
  }

  // Ensure at least one month of interest if the period is within a month boundary, as per user logic.
  conceptualMonths = conceptualMonths <= 0 ? 1 : conceptualMonths;

  const totalInterest = principal * monthlyRateDecimal * conceptualMonths;
  return totalInterest;
};

const formatCurrency = (amount) =>
  `₹${Math.round(Number(amount || 0)).toLocaleString("en-IN")}`;

const openWhatsApp = async (customer) => {
  if (!customer || (!customer.whatsapp && !customer.phone)) {
    alert("Customer's WhatsApp number is not available.");
    return;
  }

  try {
    await generateAndDownloadPDF(customer.id);
  } catch (error) {
    console.error("Failed to generate PDF before sending WhatsApp:", error);
    alert("Could not generate the PDF, but you can still send the message.");
  }

  let phone = (customer.whatsapp || customer.phone).replace(/\D/g, "");
  if (phone.length === 10) {
    phone = "91" + phone;
  }

  const { loanDetails, paymentSchedule, name } = customer;
  const totalPaid = paymentSchedule.reduce((sum, p) => sum + p.amountPaid, 0);

  // Use the same conceptual interest for displaying summary
  const totalInterest = calculateConceptualTotalInterest(
    loanDetails.principal,
    loanDetails.interestRate,
    loanDetails.firstCollectionDate,
    loanDetails.loanEndDate
  );
  const totalRepayable = loanDetails.principal + totalInterest;
  const outstanding = totalRepayable - totalPaid;

  const nextDue = paymentSchedule.find(
    (p) => p.status === "Due" || p.status === "Pending"
  );

  let message = `Hello ${name},\n\nHere is a summary of your loan with Global Finance Consultant (Finance ${
    customer.financeCount || 1
  }):\n\n`;
  message += `*Principal:* ${formatCurrency(loanDetails.principal)}\n`;
  message += `*Total Paid:* ${formatCurrency(totalPaid)}\n`;
  message += `*Outstanding Balance:* ${formatCurrency(outstanding)}\n\n`;

  if (nextDue) {
    message += `*Next Payment Due:* ${formatCurrency(
      nextDue.pendingAmount
    )} on ${nextDue.dueDate}.\n\n`;
  } else {
    message += `Your loan has been fully paid. Thank you!\n\n`;
  }
  message += `Thank you.`;

  const whatsappUrl = `https://wa.me/${phone}?text=${encodeURIComponent(
    message
  )}`;
  window.open(whatsappUrl, "_blank");
};

window.processProfitData = (allCusts) => {
  const data = [];
  const today = new Date();
  today.setHours(23, 59, 59, 999);
  [...allCusts.active, ...allCusts.settled].forEach((customer) => {
    if (customer.loanDetails && customer.paymentSchedule) {
      const totalInterest = calculateConceptualTotalInterest(
        customer.loanDetails.principal,
        customer.loanDetails.interestRate,
        customer.loanDetails.firstCollectionDate,
        customer.loanDetails.loanEndDate
      );
      const totalRepayable = customer.loanDetails.principal + totalInterest;
      if (totalRepayable === 0) return;

      const interestRatio = totalInterest / totalRepayable;

      customer.paymentSchedule.forEach((installment) => {
        if (
          installment.amountPaid > 0 &&
          new Date(installment.dueDate) <= today
        ) {
          const profitFromThisPayment = installment.amountPaid * interestRatio;
          data.push({
            date: installment.dueDate,
            profit: profitFromThisPayment,
          });
        }
      });
    }
  });
  return data;
};

document.addEventListener("DOMContentLoaded", () => {
  let recentActivities = [];
  let currentUser = null;

  const getEl = (id) => document.getElementById(id);
  const showToast = (type, title, message) => {
    const toastContainer = getEl("toast-container");
    if (!toastContainer) return;
    const toast = document.createElement("div");
    toast.className = `toast toast-${type}`;
    const icons = {
      success: "fa-check-circle",
      error: "fa-exclamation-circle",
    };
    toast.innerHTML = `<i class="fas ${icons[type]} toast-icon"></i><div><div class="toast-title">${title}</div><div class="toast-message">${message}</div></div><button class="toast-close" onclick="this.parentElement.remove()">&times;</button>`;
    toastContainer.appendChild(toast);
    setTimeout(() => {
      toast.classList.add("hide");
      setTimeout(() => toast.remove(), 400);
    }, 5000);
  };
  const toggleButtonLoading = (btn, isLoading, text = "Loading...") => {
    if (!btn) return;
    const spinner = btn.querySelector(".loading-spinner");
    const btnText = btn.querySelector("span:not(.loading-spinner)");
    if (btnText && !btnText.dataset.originalText) {
      btnText.dataset.originalText = btnText.textContent;
    }
    btn.disabled = isLoading;
    if (spinner) spinner.classList.toggle("hidden", !isLoading);
    if (btnText) {
      btnText.textContent = isLoading ? text : btnText.dataset.originalText;
    }
  };

  const calculateInstallments = (startDateStr, endDateStr, frequency) => {
    const startDate = new Date(startDateStr);
    const endDate = new Date(endDateStr);
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(0, 0, 0, 0);

    if (endDate < startDate) {
      // Use < instead of <= to allow same-day loans for daily
      throw new Error("End date must be on or after the start date.");
    }

    const diffTime = Math.abs(endDate - startDate);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;

    switch (frequency) {
      case "daily":
        return diffDays;
      case "weekly":
        return Math.ceil(diffDays / 7);
      case "monthly":
        let d1 = new Date(startDateStr);
        let d2 = new Date(endDateStr);
        let months = (d2.getFullYear() - d1.getFullYear()) * 12;
        months -= d1.getMonth();
        months += d2.getMonth();
        if (d2.getDate() < d1.getDate()) {
          months--;
        }
        return months <= 0 ? 1 : months;
      default:
        throw new Error("Invalid frequency selected.");
    }
  };

  const generateSimpleInterestSchedule = (
    principal,
    interestRate,
    n,
    firstDate,
    endDate,
    frequency
  ) => {
    const totalInterest = calculateConceptualTotalInterest(
      principal,
      interestRate,
      firstDate,
      endDate
    );
    const totalRepayable = principal + totalInterest;

    const schedule = [];
    let currentDate = new Date(firstDate);
    let cumulativeAmount = 0;

    const standardInstallment = Math.round((totalRepayable / n) * 100) / 100;

    for (let i = 1; i <= n; i++) {
      let dueDate;
      if (i === 1) {
        dueDate = new Date(currentDate);
      } else {
        if (frequency === "daily") {
          currentDate.setDate(currentDate.getDate() + 1);
        } else if (frequency === "weekly") {
          currentDate.setDate(currentDate.getDate() + 7);
        } else if (frequency === "monthly") {
          const originalDate = new Date(firstDate);
          originalDate.setMonth(originalDate.getMonth() + (i - 1));
          currentDate = originalDate;
        }
        dueDate = new Date(currentDate);
      }

      let installmentAmount;
      if (i < n) {
        installmentAmount = standardInstallment;
        cumulativeAmount += installmentAmount;
      } else {
        installmentAmount = totalRepayable - cumulativeAmount;
      }

      schedule.push({
        installment: i,
        dueDate: dueDate.toISOString().split("T")[0],
        amountDue: +installmentAmount.toFixed(2),
        amountPaid: 0,
        pendingAmount: +installmentAmount.toFixed(2),
        status: "Due",
        paidDate: null,
      });
    }
    return schedule;
  };

  const showConfirmation = (title, message, onConfirm) => {
    getEl("confirmation-title").textContent = title;
    getEl("confirmation-message").textContent = message;
    getEl("confirmation-modal").classList.add("show");
    const confirmBtn = getEl("confirmation-confirm-btn");
    const cancelBtn = getEl("confirmation-cancel-btn");
    const confirmHandler = () => {
      onConfirm();
      cleanup();
    };
    const cancelHandler = () => cleanup();
    const cleanup = () => {
      getEl("confirmation-modal").classList.remove("show");
      confirmBtn.removeEventListener("click", confirmHandler);
      cancelBtn.removeEventListener("click", cancelHandler);
    };
    confirmBtn.addEventListener("click", confirmHandler, { once: true });
    cancelBtn.addEventListener("click", cancelHandler, { once: true });
  };
  const logActivity = async (type, details) => {
    if (!currentUser) return;
    try {
      await db.collection("activities").add({
        owner_uid: currentUser.uid,
        type: type,
        details: details,
        timestamp: firebase.firestore.FieldValue.serverTimestamp(),
      });
    } catch (error) {
      console.error("Failed to log activity:", error);
    }
  };

  const calculateKeyStats = (activeLoans, settledLoans) => {
    let totalPrincipal = 0,
      totalOutstanding = 0,
      totalInterestEarned = 0;

    [...activeLoans, ...settledLoans].forEach((c) => {
      if (c.loanDetails && c.paymentSchedule) {
        totalPrincipal += c.loanDetails.principal;

        const totalInterest = calculateConceptualTotalInterest(
          c.loanDetails.principal,
          c.loanDetails.interestRate,
          c.loanDetails.firstCollectionDate,
          c.loanDetails.loanEndDate
        );
        const totalRepayable = c.loanDetails.principal + totalInterest;
        const totalPaid = c.paymentSchedule.reduce(
          (sum, p) => sum + p.amountPaid,
          0
        );

        const interestPaid = Math.max(0, totalPaid - c.loanDetails.principal);
        totalInterestEarned += interestPaid;

        if (c.status === "active") {
          totalOutstanding += totalRepayable - totalPaid;
        }
      }
    });

    return {
      activeLoanCount: activeLoans.filter((c) => c.loanDetails).length,
      settledLoanCount: settledLoans.filter((c) => c.loanDetails).length,
      totalPrincipal,
      totalOutstanding,
      totalInterest: totalInterestEarned,
    };
  };

  const updateSidebarStats = (stats) => {
    getEl("sidebar-active-loans").textContent = stats.activeLoanCount;
    getEl("sidebar-settled-loans").textContent = stats.settledLoanCount;
    getEl("sidebar-interest-earned").textContent = formatCurrency(
      stats.totalInterest
    );
    getEl("sidebar-outstanding").textContent = formatCurrency(
      stats.totalOutstanding
    );
  };

  const loadAndRenderAll = async () => {
    if (!currentUser) return;
    try {
      const customerQuery = db
        .collection("customers")
        .where("owner", "==", currentUser.uid);
      const [customerSnapshot, activitiesSnapshot] = await Promise.all([
        customerQuery.orderBy("createdAt", "desc").get(),
        db
          .collection("activities")
          .where("owner_uid", "==", currentUser.uid)
          .orderBy("timestamp", "desc")
          .limit(10)
          .get(),
      ]);

      const allDocs = customerSnapshot.docs.map((doc) => ({
        id: doc.id,
        ...doc.data(),
      }));

      allDocs.forEach((doc) => {
        if (doc.loanDetails && !doc.loanDetails.loanEndDate) {
          const schedule = doc.paymentSchedule;
          if (schedule && schedule.length > 0) {
            doc.loanDetails.loanEndDate = schedule[schedule.length - 1].dueDate;
          }
        }
      });

      window.allCustomers.active = allDocs.filter((c) => c.status === "active");
      window.allCustomers.settled = allDocs.filter(
        (c) => c.status === "settled" || c.status === "Refinanced"
      );

      recentActivities = activitiesSnapshot.docs.map((doc) => ({
        id: doc.id,
        ...doc.data(),
      }));

      const stats = calculateKeyStats(
        window.allCustomers.active,
        window.allCustomers.settled
      );
      populatePageContent(stats);
      updateSidebarStats(stats);

      const profitData = window.processProfitData(window.allCustomers);
      if (typeof renderDashboardCharts === "function") {
        renderDashboardCharts(
          window.allCustomers.active,
          window.allCustomers.settled,
          profitData
        );
      }
    } catch (error) {
      console.error("Error loading data:", error);
      showToast("error", "Load Failed", "Could not fetch all required data.");
    }
  };

  const renderActivityLog = () => {
    const container = getEl("recent-activity-container");
    if (!container) return;
    if (!recentActivities || recentActivities.length === 0) {
      container.innerHTML = `<ul class="activity-list"><li class="activity-item" style="cursor:default;">No recent activities found.</li></ul>`;
      return;
    }
    const activityHTML = recentActivities
      .map((act) => {
        let icon = "fa-info-circle";
        let text = "New activity";
        const customerName = act.details.customerName
          ? `<strong>${act.details.customerName}</strong>`
          : "";
        const amount = act.details.amount
          ? formatCurrency(act.details.amount)
          : "";
        switch (act.type) {
          case "NEW_LOAN":
            icon = "fa-user-plus text-success";
            text = `New loan for ${customerName} of ${amount}`;
            break;
          case "PAYMENT_RECEIVED":
            icon = "fa-check-circle text-success";
            text = `Payment of ${amount} received from ${customerName}`;
            break;
          case "LOAN_SETTLED":
            icon = "fa-flag-checkered text-primary";
            text = `Loan settled for ${customerName}`;
            break;
        }
        const timestamp = act.timestamp
          ? new Date(act.timestamp.seconds * 1000).toLocaleString()
          : "Just now";
        return `<li class="activity-item" id="activity-${act.id}"><div class="activity-info"><i class="fas ${icon} activity-icon"></i><div class="activity-text"><span class="activity-name">${text}</span><span class="activity-date">${timestamp}</span></div></div><div class="activity-actions"><button class="delete-activity-btn" data-id="${act.id}" title="Delete Activity"><i class="fas fa-trash-alt"></i></button></div></li>`;
      })
      .join("");
    container.innerHTML = `<ul class="activity-list">${activityHTML}</ul>`;
  };

  const populatePageContent = (stats) => {
    getEl(
      "dashboard-section"
    ).innerHTML = `<div class="stats-container"><div class="stat-card"><div class="stat-title">Total Principal</div><div class="stat-value">${formatCurrency(
      stats.totalPrincipal
    )}</div></div><div class="stat-card"><div class="stat-title">Outstanding</div><div class="stat-value">${formatCurrency(
      stats.totalOutstanding
    )}</div></div><div class="stat-card"><div class="stat-title">Interest Earned</div><div class="stat-value">${formatCurrency(
      stats.totalInterest
    )}</div></div><div class="stat-card"><div class="stat-title">Active Loans</div><div class="stat-value">${
      stats.activeLoanCount
    }</div></div></div><div class="dashboard-grid"><div class="form-card chart-card"><h3 >Portfolio Overview</h3><div class="chart-container"><canvas id="portfolioChart"></canvas></div></div><div class="form-card chart-card grid-col-span-2"><h3 >Profit Over Time <div class="chart-controls" id="profit-chart-controls"><button class="btn btn-sm btn-outline active" data-frame="monthly">Month</button><button class="btn btn-sm btn-outline" data-frame="yearly">Year</button></div></h3><div class="chart-container"><canvas id="profitChart"></canvas></div></div><div class="form-card"><h3><i class="fas fa-clock" style="color:var(--primary)"></i> Upcoming Installments</h3><div id="upcoming-emi-container" class="activity-container"></div></div><div class="form-card"><h3><i class="fas fa-exclamation-triangle" style="color:var(--danger)"></i> Overdue Installments</h3><div id="overdue-emi-container" class="activity-container"></div></div><div class="form-card"><h3><i class="fas fa-history"></i> Recent Activity <button class="btn btn-danger btn-sm" id="clear-all-activities-btn" title="Clear all activities"><i class="fas fa-trash"></i></button></h3><div id="recent-activity-container" class="activity-container"></div></div></div>`;
    getEl(
      "calculator-section"
    ).innerHTML = `<div class="form-card"><h3><i class="fas fa-calculator"></i> Simple Interest Loan Calculator</h3><form id="emi-calculator-form"><div class="form-group"><label for="calc-principal">Loan Amount (₹)</label><input type="number" id="calc-principal" class="form-control" placeholder="e.g., 50000" required /></div><div class="form-row" style="grid-template-columns: 1fr 1fr 1fr;"><div class="form-group"><label for="calc-rate">Monthly Interest Rate (%)</label><input type="number" id="calc-rate" class="form-control" placeholder="e.g., 10" step="0.01" required /></div><div class="form-group"><label for="collection-frequency-calc">Collection Frequency</label><select id="collection-frequency-calc" class="form-control" required><option value="daily">Daily</option><option value="weekly">Weekly</option><option value="monthly" selected>Monthly</option></select></div><div class="form-group"><label for="calc-tenure">Number of Installments</label><input type="number" id="calc-tenure" class="form-control" placeholder="e.g., 12" required /></div></div><button type="submit" class="btn btn-primary">Calculate</button></form><div id="calculator-results" class="hidden" style="margin-top: 2rem; border-top: 1px solid var(--border-color); padding-top: 1.5rem;"><h4>Calculation Result</h4><div class="calc-result-item"><span>Per Installment Amount</span><span id="result-emi"></span></div><div class="calc-result-item"><span>Total Interest</span><span id="result-interest"></span></div><div class="calc-result-item"><span>Total Payment</span><span id="result-total"></span></div></div></div>`;
    getEl("settings-section").innerHTML = `<div class="settings-grid">
        <div class="form-card setting-card"><div class="setting-card-header"><i class="fas fa-shield-alt"></i><h3>Security</h3></div><div class="setting-card-body"><p class="setting-description">Manage your account security settings.</p><form id="change-password-form"><div class="form-group"><label for="current-password">Current Password</label><input type="password" id="current-password" class="form-control" required/></div><div class="form-row"><div class="form-group"><label for="new-password">New Password</label><input type="password" id="new-password" class="form-control" required/></div><div class="form-group"><label for="confirm-password">Confirm New Password</label><input type="password" id="confirm-password" class="form-control" required/></div></div><button id="change-password-btn" type="submit" class="btn btn-primary"><span class="loading-spinner hidden"></span><span>Update Password</span></button></form></div></div>
        <div class="form-card setting-card"><div class="setting-card-header"><i class="fas fa-palette"></i><h3>Appearance</h3></div><div class="setting-card-body"><p class="setting-description">Customize the look and feel of the application.</p><div class="setting-item"><div class="setting-label"><i class="fas fa-moon"></i><span>Dark Mode</span></div><div class="setting-control"><label class="switch"><input type="checkbox" id="dark-mode-toggle" /><span class="slider round"></span></label></div></div></div></div>
        <div class="form-card setting-card"><div class="setting-card-header"><i class="fas fa-database"></i><h3>Data Management</h3></div><div class="setting-card-body"><p class="setting-description">Backup your data or restore from a file.</p><div class="setting-item"><div class="setting-label"><i class="fas fa-download"></i><span>Export Data</span></div><div class="setting-control"><button class="btn btn-outline" id="export-backup-btn">Download Backup</button></div></div><hr class="form-hr"><div class="setting-item column"><div class="setting-label"><i class="fas fa-upload"></i><span>Import Data</span></div><div class="import-controls"><input type="file" id="import-backup-input" accept=".json" class="hidden"><label for="import-backup-input" class="btn btn-outline"><i class="fas fa-file-import"></i> Choose File</label><span id="file-name-display" class="file-name">No file chosen</span></div><p class="warning-text"><i class="fas fa-exclamation-triangle"></i> This will overwrite all current data.</p><button class="btn btn-danger" id="import-backup-btn">Import & Overwrite</button></div></div></div>
        <div class="form-card setting-card"><div class="setting-card-header"><i class="fas fa-user-cog"></i><h3>Account</h3></div><div class="setting-card-body"><p class="setting-description">Log out from your current session.</p><div class="setting-item"><div class="setting-label"><i class="fas fa-sign-out-alt"></i><span>Logout</span></div><div class="setting-control"><button class="btn btn-outline" id="logout-settings-btn">Logout Now</button></div></div></div></div>
    </div>`;

    populateCustomerLists();
    populateTodaysCollection();
    renderUpcomingAndOverdueEmis(window.allCustomers.active);
    renderActivityLog();
  };

  const renderIndividualLoanList = (element, data) => {
    if (!element) return;
    element.innerHTML = "";
    if (data.length === 0) {
      element.innerHTML = `<li class="activity-item" style="cursor:default; justify-content:center;"><p>No customers found.</p></li>`;
      return;
    }
    data.forEach((c) => {
      const li = document.createElement("li");
      li.className = "customer-item";
      if (!c.loanDetails || !c.paymentSchedule) {
        li.innerHTML = `<div class="customer-info" data-id="${c.id}"><div class="customer-name">${c.name}</div><div class="customer-details"><span>Data format error</span></div></div><div class="customer-actions"><span class="view-details-prompt" data-id="${c.id}">View Details</span></div>`;
        element.appendChild(li);
        return;
      }

      const financeCount = c.financeCount || 1;
      const totalPaid = c.paymentSchedule.reduce(
        (sum, p) => sum + p.amountPaid,
        0
      );
      const interestEarned = Math.max(0, totalPaid - c.loanDetails.principal);

      const financeStatusText = `Finance ${financeCount} - ${c.status}`;
      const nameHtml = `<div class="customer-name">${c.name} <span class="finance-status-badge">${financeStatusText}</span></div>`;
      const detailsHtml = `<span>Principal: ${formatCurrency(
        c.loanDetails.principal
      )}</span><span class="list-profit-display success">Interest: ${formatCurrency(
        interestEarned
      )}</span>`;

      const deleteButton = `<button class="btn btn-danger btn-sm delete-customer-btn" data-id="${c.id}" title="Delete Customer"><i class="fas fa-trash-alt"></i></button>`;

      li.innerHTML = `<div class="customer-info" data-id="${c.id}">${nameHtml}<div class="customer-details">${detailsHtml}</div></div><div class="customer-actions">${deleteButton}<span class="view-details-prompt" data-id="${c.id}">View Details</span></div>`;
      element.appendChild(li);
    });
  };

  const renderActiveCustomerList = (customerArray) => {
    const listEl = getEl("customers-list");
    if (!listEl) return;

    const customerGroups = new Map();
    customerArray.forEach((customer) => {
      if (!customerGroups.has(customer.name)) {
        customerGroups.set(customer.name, []);
      }
      customerGroups.get(customer.name).push(customer);
    });

    if (customerGroups.size === 0) {
      listEl.innerHTML = `<li class="activity-item" style="cursor:default; justify-content:center;"><p>No customers found.</p></li>`;
      return;
    }

    let listHtml = "";
    for (const [name, loans] of customerGroups.entries()) {
      const latestLoan = [...loans].sort(
        (a, b) => (b.financeCount || 1) - (a.financeCount || 1)
      )[0];
      const totalActiveLoans = loans.length;
      const totalOutstanding = loans.reduce((sum, loan) => {
        const totalInterest = calculateConceptualTotalInterest(
          loan.loanDetails.principal,
          loan.loanDetails.interestRate,
          loan.loanDetails.firstCollectionDate,
          loan.loanDetails.loanEndDate
        );
        const totalRepayable = loan.loanDetails.principal + totalInterest;
        const totalPaid = loan.paymentSchedule.reduce(
          (s, p) => s + p.amountPaid,
          0
        );
        return sum + (totalRepayable - totalPaid);
      }, 0);

      const nameHtml = `<div class="customer-name">${name}</div>`;
      const detailsHtml = `<span>Total Outstanding: ${formatCurrency(
        totalOutstanding
      )}</span>`;
      const loanCountBadge =
        totalActiveLoans > 1
          ? `<span class="finance-count-badge">${totalActiveLoans}</span>`
          : "";

      listHtml += `
            <li class="customer-item">
                <div class="customer-info" data-id="${latestLoan.id}">
                    ${nameHtml}
                    <div class="customer-details">${detailsHtml}</div>
                </div>
                <div class="customer-actions">
                    ${loanCountBadge}
                    <span class="view-details-prompt" data-id="${latestLoan.id}">View Details</span>
                </div>
            </li>`;
    }
    listEl.innerHTML = listHtml;
  };

  const populateCustomerLists = () => {
    getEl(
      "active-accounts-section"
    ).innerHTML = `<div class="form-card"><div class="card-header"><h3>Active Accounts</h3><button class="btn btn-outline" id="export-active-btn"><i class="fas fa-file-excel"></i> Export to Excel</button></div><div class="form-group"><input type="text" id="search-customers" class="form-control" placeholder="Search active customers..." /></div><ul id="customers-list" class="customer-list"></ul></div>`;
    renderActiveCustomerList(window.allCustomers.active);

    getEl(
      "settled-accounts-section"
    ).innerHTML = `<div class="form-card"><div class="card-header"><h3>Settled Accounts (${window.allCustomers.settled.length})</h3><div style="display: flex; gap: 0.75rem;"><button class="btn btn-danger" id="delete-all-settled-btn"><i class="fas fa-trash-alt"></i> Delete All</button><button class="btn btn-outline" id="export-settled-btn"><i class="fas fa-file-excel"></i> Export to Excel</button></div></div><ul id="settled-customers-list" class="customer-list"></ul></div>`;
    renderIndividualLoanList(
      getEl("settled-customers-list"),
      window.allCustomers.settled
    );
  };

  const renderUpcomingAndOverdueEmis = (activeLoans) => {
    const upcomingContainer = getEl("upcoming-emi-container");
    const overdueContainer = getEl("overdue-emi-container");
    if (!upcomingContainer || !overdueContainer) return;
    let upcoming = [],
      overdue = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    activeLoans
      .filter((c) => c.loanDetails && c.paymentSchedule)
      .forEach((customer) => {
        const nextDueInstallment = customer.paymentSchedule.find(
          (p) => p.status === "Due" || p.status === "Pending"
        );
        if (nextDueInstallment) {
          const dueDate = new Date(nextDueInstallment.dueDate);
          const entry = {
            name: `${customer.name} (F${customer.financeCount || 1})`,
            amount: nextDueInstallment.pendingAmount,
            date: nextDueInstallment.dueDate,
            customerId: customer.id,
          };
          if (dueDate < today) overdue.push(entry);
          else upcoming.push(entry);
        }
      });

    upcoming.sort((a, b) => new Date(a.date) - new Date(b.date));
    overdue.sort((a, b) => new Date(a.date) - new Date(b.date));

    const renderEmiList = (items) => {
      if (items.length === 0)
        return `<ul class="activity-list"><li class="activity-item" style="cursor:default; justify-content:center;">No items found.</li></ul>`;
      return `<ul class="activity-list">${items
        .slice(0, 5)
        .map(
          (item) =>
            `<li class="activity-item" data-id="${
              item.customerId
            }"><div class="activity-info"><span class="activity-name">${
              item.name
            }</span><span class="activity-date">${
              item.date
            }</span></div><div class="activity-value"><span class="activity-amount">${formatCurrency(
              item.amount
            )}</span></div></li>`
        )
        .join("")}</ul>`;
    };
    upcomingContainer.innerHTML = renderEmiList(upcoming);
    overdueContainer.innerHTML = renderEmiList(overdue);
  };

  const populateTodaysCollection = () => {
    const container = getEl("todays-collection-section");
    if (!container) return;
    const today = new Date().toISOString().split("T")[0];
    const dueToday = window.allCustomers.active
      .map((cust) => {
        const dueInstallment = cust.paymentSchedule?.find(
          (inst) =>
            inst.dueDate === today &&
            (inst.status === "Due" || inst.status === "Pending")
        );
        return dueInstallment ? { ...cust, dueInstallment } : null;
      })
      .filter(Boolean);

    const totalDue = dueToday.reduce(
      (sum, item) => sum + item.dueInstallment.pendingAmount,
      0
    );
    let listHtml = "";
    if (dueToday.length > 0) {
      listHtml = dueToday
        .map(
          (item) =>
            `<li class="activity-item" data-id="${
              item.id
            }"><div class="activity-info"><span class="activity-name">${
              item.name
            } (F${
              item.financeCount || 1
            })</span><span class="activity-details">Installment #${
              item.dueInstallment.installment
            }</span></div><div class="activity-value"><span class="activity-amount">${formatCurrency(
              item.dueInstallment.pendingAmount
            )}</span></div></li>`
        )
        .join("");
    } else {
      listHtml = `<li class="activity-item" style="cursor:default; justify-content:center;">No collections due today.</li>`;
    }
    container.innerHTML = `<div class="form-card"><div class="card-header"><h3>Due Today (${
      dueToday.length
    })</h3><div class="stat-card" style="padding: 0.5rem 1rem; text-align: right;"><div class="stat-title">Total Due Today</div><div class="stat-value" style="font-size: 1.5rem;">${formatCurrency(
      totalDue
    )}</div></div></div><ul class="activity-list">${listHtml}</ul></div>`;
  };

  const showCustomerDetails = (customerId) => {
    const customer = [
      ...window.allCustomers.active,
      ...window.allCustomers.settled,
    ].find((c) => c.id === customerId);
    if (!customer) return;

    const allLoansForCustomer = [
      ...window.allCustomers.active,
      ...window.allCustomers.settled,
    ]
      .filter((c) => c.name === customer.name)
      .sort((a, b) => (a.financeCount || 1) - (b.financeCount || 1));

    const loansToDisplay = [
      ...new Map(allLoansForCustomer.map((item) => [item.id, item])).values(),
    ];

    const switcherContainer = getEl("loan-switcher-container");
    const optionsContainer = switcherContainer.querySelector(".custom-options");
    const triggerText = switcherContainer.querySelector(
      ".custom-select-trigger span"
    );

    if (loansToDisplay.length > 1) {
      switcherContainer.classList.remove("hidden");
      optionsContainer.innerHTML = loansToDisplay
        .map(
          (loan) =>
            `<div class="custom-option ${
              loan.id === customerId ? "selected" : ""
            }" data-value="${loan.id}">
                Finance ${loan.financeCount || 1}
            </div>`
        )
        .join("");
      const currentLoan =
        loansToDisplay.find((l) => l.id === customerId) || customer;
      triggerText.textContent = `Finance ${currentLoan.financeCount || 1}`;
    } else {
      switcherContainer.classList.add("hidden");
    }

    renderLoanDetails(customerId);
    getEl("customer-details-modal").classList.add("show");
  };

  const renderLoanDetails = (customerId) => {
    const customer = [
      ...window.allCustomers.active,
      ...window.allCustomers.settled,
    ].find((c) => c.id === customerId);
    if (!customer) return;

    const modalBody = getEl("details-modal-body");
    getEl("details-modal-title").textContent = `Details: ${customer.name}`;
    const frequencyBadge = getEl("details-modal-frequency");
    if (frequencyBadge && customer.loanDetails?.frequency) {
      frequencyBadge.textContent = customer.loanDetails.frequency;
      frequencyBadge.className = `loan-frequency-badge frequency-${customer.loanDetails.frequency}`;
    }

    getEl("generate-pdf-btn").dataset.id = customerId;
    getEl("send-whatsapp-btn").dataset.id = customerId;

    if (!customer.loanDetails || !customer.paymentSchedule) {
      modalBody.innerHTML = `<p style="padding: 2rem; text-align: center;">This customer has incomplete or outdated loan data.</p>`;
      return;
    }

    const { paymentSchedule: schedule, loanDetails: details } = customer;

    const totalInterest = calculateConceptualTotalInterest(
      details.principal,
      details.interestRate,
      details.firstCollectionDate,
      details.loanEndDate
    );
    const totalPaid = schedule.reduce((sum, p) => sum + p.amountPaid, 0);
    const totalRepayable = details.principal + totalInterest;
    const remainingToCollect = totalRepayable - totalPaid;
    const paidInstallments = schedule.filter((p) => p.status === "Paid").length;
    const progress =
      schedule.length > 0 ? (paidInstallments / schedule.length) * 100 : 0;
    const totalInterestPaid = Math.max(0, totalPaid - details.principal);
    const nextDue = schedule.find(
      (p) => p.status === "Due" || p.status === "Pending"
    );

    let actionButtons = "";
    if (customer.status === "active") {
      actionButtons += `<button class="btn btn-primary" id="add-new-loan-btn" data-id="${customer.id}"><i class="fas fa-plus-circle"></i> Add New Loan</button>`;
      actionButtons += `<button class="btn btn-success" id="settle-loan-btn" data-id="${customer.id}"><i class="fas fa-check-circle"></i> Settle Loan</button>`;
    }

    const createKycViewButton = (url, label = "View File") =>
      url
        ? `<a href="${url}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-outline">${label}</a>`
        : '<span class="value">N/A</span>';
    const kycDocs = customer.kycDocs || {};
    const aadharButton = createKycViewButton(kycDocs.aadharUrl);
    const panButton = createKycViewButton(kycDocs.panUrl);
    const picButton = createKycViewButton(kycDocs.picUrl, "View Photo");
    const bankButton = createKycViewButton(kycDocs.bankDetailsUrl);

    modalBody.innerHTML = `<div class="details-view-grid"><div class="customer-profile-panel"><div class="profile-header"><div class="profile-avatar">${customer.name
      .charAt(0)
      .toUpperCase()}</div><h3 class="profile-name">${
      customer.name
    }</h3><p class="profile-contact">${
      customer.phone || "N/A"
    }</p></div><div class="profile-section"><h4>Personal Details</h4><div class="profile-stat"><span class="label">Father's Name</span><span class="value">${
      customer.fatherName || "N/A"
    }</span></div><div class="profile-stat"><span class="label">Address</span><span class="value">${
      customer.address || "N/A"
    }</span></div></div><div class="profile-section"><h4>KYC Documents</h4><div class="profile-stat"><span class="label">Aadhar Card</span>${aadharButton}</div><div class="profile-stat"><span class="label">PAN Card</span>${panButton}</div><div class="profile-stat"><span class="label">Client Photo</span>${picButton}</div><div class="profile-stat"><span class="label">Bank Details</span>${bankButton}</div></div><div class="loan-progress-section"><h4>Loan Progress (${paidInstallments} of ${
      schedule.length
    } Paid)</h4><div class="progress-bar"><div class="progress-bar-inner" style="width: ${progress}%;"></div></div></div><div class="loan-actions"><button class="btn btn-outline" id="edit-customer-info-btn" data-id="${
      customer.id
    }"><i class="fas fa-edit"></i> Edit Info</button>${actionButtons}</div></div><div class="emi-schedule-panel"><div class="emi-table-container"><table class="emi-table"><thead><tr><th>#</th><th>Due Date</th><th>Amount Due</th><th>Amount Paid</th><th>Status</th><th class="no-pdf">Action</th></tr></thead><tbody id="emi-schedule-body-details"></tbody></table></div><div class="loan-summary-box"><h4>Loan Summary</h4><div class="calc-result-item"><span>Principal Amount</span><span>${formatCurrency(
      details.principal
    )}</span></div><div class="calc-result-item"><span>Total Interest Paid</span><span>${formatCurrency(
      totalInterestPaid
    )}</span></div><div class="calc-result-item"><span>Outstanding Amount</span><span>${formatCurrency(
      remainingToCollect
    )}</span></div><div class="calc-result-item"><span>Next Due</span><span>${
      nextDue ? nextDue.dueDate : "N/A"
    }</span></div></div><div class="modal-summary-stats"><div class="summary-stat-item"><span class="label">Amount Received</span><span class="value received">${formatCurrency(
      totalPaid
    )}</span></div><div class="summary-stat-item"><span class="label">Amount Remaining</span><span class="value remaining">${formatCurrency(
      remainingToCollect
    )}</span></div></div></div></div>`;

    const emiTableBody = modalBody.querySelector("#emi-schedule-body-details");
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    schedule.forEach((inst) => {
      const tr = document.createElement("tr");
      if (inst.status === "Paid") tr.classList.add("emi-paid-row");
      const isOverdue =
        new Date(inst.dueDate) < today &&
        (inst.status === "Due" || inst.status === "Pending");
      let statusClass = inst.status.toLowerCase();
      let statusText = inst.status.toUpperCase();
      if (isOverdue) {
        statusClass = "overdue";
        statusText = "OVERDUE";
      }
      if (inst.status === "Pending") {
        statusText += ` (${formatCurrency(inst.pendingAmount)} due)`;
      }
      let actionButtons = "";
      if (
        customer.status === "active" &&
        (inst.status === "Due" || inst.status === "Pending")
      ) {
        actionButtons += `<button class="btn btn-success btn-sm record-payment-btn" data-installment="${inst.installment}" data-id="${customer.id}">Pay</button>`;
      }
      if (customer.status === "active" && inst.amountPaid > 0) {
        actionButtons += `<button class="btn btn-outline btn-sm record-payment-btn" data-installment="${inst.installment}" data-id="${customer.id}">Edit</button>`;
      }
      tr.innerHTML = `<td>${inst.installment}</td><td>${
        inst.dueDate
      }</td><td>${formatCurrency(inst.amountDue)}</td><td>${formatCurrency(
        inst.amountPaid
      )}</td><td><span class="emi-status status-${statusClass}">${statusText}</span></td><td class="no-pdf">${actionButtons}</td>`;
      emiTableBody.appendChild(tr);
    });
  };

  async function settleLoanById(loanId) {
    const customerToSettle = window.allCustomers.active.find(
      (c) => c.id === loanId
    );
    if (!customerToSettle) {
      showToast("error", "Not Found", "The selected loan could not be found.");
      return;
    }
    try {
      await db
        .collection("customers")
        .doc(loanId)
        .update({ status: "settled" });
      await logActivity("LOAN_SETTLED", {
        customerName: customerToSettle.name,
        financeCount: customerToSettle.financeCount || 1,
      });
      showToast(
        "success",
        "Loan Settled",
        `Finance ${customerToSettle.financeCount || 1} moved to settled.`
      );
      getEl("customer-details-modal").classList.remove("show");
      getEl("settle-selection-modal").classList.remove("show");
      await loadAndRenderAll();
    } catch (error) {
      showToast("error", "Settle Failed", error.message);
    }
  }

  auth.onAuthStateChanged(async (user) => {
    currentUser = user;
    if (user) {
      getEl("auth-container").classList.add("hidden");
      getEl("admin-dashboard").classList.remove("hidden");
      await loadAndRenderAll();
      if (window.applyTheme) {
        const savedTheme = localStorage.getItem("theme") || "default";
        window.applyTheme(savedTheme);
      }
    } else {
      window.allCustomers = { active: [], settled: [] };
      recentActivities = [];
      getEl("auth-container").classList.remove("hidden");
      getEl("admin-dashboard").classList.add("hidden");
    }
  });

  function updateInstallmentPreview() {
    const previewDiv = getEl("installment-preview");
    const p = parseFloat(getEl("principal-amount").value);
    const r = parseFloat(getEl("interest-rate-modal").value);
    const freq = getEl("collection-frequency").value;
    const firstDate = getEl("first-collection-date").value;
    const endDate = getEl("loan-end-date").value;

    if (!p || !r || !firstDate || !endDate) {
      previewDiv.classList.add("hidden");
      return;
    }

    try {
      const n = calculateInstallments(firstDate, endDate, freq);
      if (n <= 0) {
        throw new Error("Invalid date range for the selected frequency.");
      }
      const totalInterest = calculateConceptualTotalInterest(
        p,
        r,
        firstDate,
        endDate
      );
      const totalRepayable = p + totalInterest;
      const installmentAmount = totalRepayable / n;

      previewDiv.innerHTML = `<p>Calculated: <strong>${n} installments</strong> of <strong>${formatCurrency(
        installmentAmount
      )}</strong> each.</p>`;
      previewDiv.classList.remove("error", "hidden");
    } catch (error) {
      previewDiv.innerHTML = `<p>${error.message}</p>`;
      previewDiv.classList.add("error");
      previewDiv.classList.remove("hidden");
    }
  }

  function updateNewLoanInstallmentPreview() {
    const previewDiv = getEl("new-loan-installment-preview");
    const p = parseFloat(getEl("new-loan-principal").value);
    const r = parseFloat(getEl("new-loan-interest-rate").value);
    const freq = getEl("new-loan-frequency").value;
    const firstDate = getEl("new-loan-start-date").value;
    const endDate = getEl("new-loan-end-date").value;

    if (!p || !r || !firstDate || !endDate) {
      previewDiv.classList.add("hidden");
      return;
    }

    try {
      const n = calculateInstallments(firstDate, endDate, freq);
      if (n <= 0) {
        throw new Error("Invalid date range for the selected frequency.");
      }
      const totalInterest = calculateConceptualTotalInterest(
        p,
        r,
        firstDate,
        endDate
      );
      const totalRepayable = p + totalInterest;
      const installmentAmount = totalRepayable / n;

      previewDiv.innerHTML = `<p>New Loan: <strong>${n} installments</strong> of <strong>${formatCurrency(
        installmentAmount
      )}</strong> each.</p>`;
      previewDiv.classList.remove("error", "hidden");
    } catch (error) {
      previewDiv.innerHTML = `<p>${error.message}</p>`;
      previewDiv.classList.add("error");
      previewDiv.classList.remove("hidden");
    }
  }

  function setAutomaticFirstDate() {
    const freq = getEl("collection-frequency").value;
    const firstDateInput = getEl("first-collection-date");
    const today = new Date();

    switch (freq) {
      case "daily":
        break;
      case "weekly":
        today.setDate(today.getDate() + 7);
        break;
      case "monthly":
        today.setMonth(today.getMonth() + 1);
        break;
    }
    firstDateInput.value = today.toISOString().split("T")[0];
    firstDateInput.dispatchEvent(new Event("change"));
  }

  // NEW: Function to set automatic first date for the "Add New Loan" modal
  function setAutomaticNewLoanFirstDate() {
    const freq = getEl("new-loan-frequency").value;
    const firstDateInput = getEl("new-loan-start-date");
    const today = new Date();

    switch (freq) {
      case "daily":
        break;
      case "weekly":
        today.setDate(today.getDate() + 7);
        break;
      case "monthly":
        today.setMonth(today.getMonth() + 1);
        break;
    }
    firstDateInput.value = today.toISOString().split("T")[0];
    firstDateInput.dispatchEvent(new Event("change"));
  }

  function initializeEventListeners() {
    getEl("login-form")?.addEventListener("submit", async (e) => {
      e.preventDefault();
      const btn = getEl("login-btn");
      toggleButtonLoading(btn, true, "Logging In...");
      try {
        await auth.signInWithEmailAndPassword(
          getEl("login-email").value,
          getEl("login-password").value
        );
      } catch (error) {
        showToast("error", "Login Failed", error.message);
        toggleButtonLoading(btn, false);
      }
    });

    document.body.addEventListener("click", async (e) => {
      const target = e.target;
      const button = target.closest("button");

      const customSelect = target.closest(".custom-select");
      if (customSelect && !target.closest(".custom-option")) {
        customSelect.classList.toggle("open");
      } else if (!target.closest(".custom-select-wrapper")) {
        document
          .querySelectorAll(".custom-select.open")
          .forEach((select) => select.classList.remove("open"));
      }

      const customOption = target.closest(".custom-option");
      if (customOption) {
        const selectWrapper = target.closest(".custom-select-wrapper");
        const triggerText = selectWrapper.querySelector(
          ".custom-select-trigger span"
        );
        triggerText.textContent = customOption.textContent.trim();
        selectWrapper.querySelector(".custom-select").classList.remove("open");

        const options = selectWrapper.querySelectorAll(".custom-option");
        options.forEach((opt) => opt.classList.remove("selected"));
        customOption.classList.add("selected");

        renderLoanDetails(customOption.dataset.value);
      }

      if (target.id === "forgot-password-link") {
        e.preventDefault();
        const email = prompt(
          "Please enter your email to receive a password reset link:"
        );
        if (email) {
          try {
            await auth.sendPasswordResetEmail(email);
            showToast(
              "success",
              "Email Sent",
              "Password reset link sent to your email."
            );
          } catch (error) {
            showToast("error", "Error", error.message);
          }
        }
      }
      if (
        target.closest("#mobile-menu-btn") ||
        target.closest("#sidebar-overlay")
      ) {
        getEl("sidebar")?.classList.toggle("show");
        getEl("sidebar-overlay")?.classList.toggle("show");
        getEl("mobile-menu-btn")?.classList.toggle("is-hidden");
      }

      const menuItem = target.closest(".menu-item:not(#change-theme-btn)");
      if (menuItem) {
        e.preventDefault();
        document
          .querySelectorAll(".menu-item")
          .forEach((i) => i.classList.remove("active"));
        menuItem.classList.add("active");
        const sectionId = `${menuItem.dataset.section}-section`;
        document
          .querySelectorAll(".section-content")
          .forEach((s) => s.classList.remove("is-active"));
        getEl(sectionId)?.classList.add("is-active");
        getEl("section-title").textContent =
          menuItem.querySelector("span").textContent;
        if (getEl("sidebar")?.classList.contains("show")) {
          getEl("sidebar")?.classList.remove("show");
          getEl("sidebar-overlay")?.classList.remove("show");
        }
      }
      const customerItemInfo = target.closest(
        ".customer-info, .view-details-prompt, .activity-item[data-id]"
      );
      if (customerItemInfo && customerItemInfo.dataset.id) {
        showCustomerDetails(customerItemInfo.dataset.id);
      }
      if (target.closest(".modal-close, [data-close-modal]")) {
        target.closest(".modal").classList.remove("show");
      }
      if (button) {
        if (button.id === "generate-pdf-btn") {
          const customerId = button.dataset.id;
          if (customerId) generateAndDownloadPDF(customerId);
        } else if (button.id === "send-whatsapp-btn") {
          const customerId = button.dataset.id;
          const customer = [
            ...window.allCustomers.active,
            ...window.allCustomers.settled,
          ].find((c) => c.id === customerId);
          if (customer) {
            openWhatsApp(customer);
          }
        } else if (button.classList.contains("record-payment-btn")) {
          const customerId = button.dataset.id;
          const installmentNum = parseInt(button.dataset.installment, 10);
          const customer = window.allCustomers.active.find(
            (c) => c.id === customerId
          );
          if (!customer) return;
          const installment = customer.paymentSchedule.find(
            (p) => p.installment === installmentNum
          );

          getEl("payment-customer-id").value = customerId;
          getEl("payment-installment-number").value = installmentNum;

          const modal = getEl("payment-modal");
          modal.querySelector(".payment-customer-name").textContent =
            customer.name;
          modal.querySelector(".payment-customer-avatar").textContent =
            customer.name.charAt(0).toUpperCase();
          getEl("payment-installment-display").textContent = installmentNum;
          getEl("payment-due-display").textContent = formatCurrency(
            installment.amountDue
          );

          const paymentAmountInput = getEl("payment-amount");
          paymentAmountInput.value =
            installment.amountPaid > 0 ? installment.amountPaid : "";
          paymentAmountInput.setAttribute("max", installment.amountDue);

          const updatePendingDisplay = () => {
            const amountPaid = parseFloat(paymentAmountInput.value) || 0;
            const pendingAmount = installment.amountDue - amountPaid;
            const pendingContainer = modal.querySelector(
              ".payment-pending-display"
            );
            if (pendingAmount > 0.001) {
              modal.querySelector(".payment-pending-value").textContent =
                formatCurrency(pendingAmount);
              pendingContainer.style.display = "block";
            } else {
              pendingContainer.style.display = "none";
            }
          };

          paymentAmountInput.oninput = updatePendingDisplay;

          const payFullBtn = getEl("pay-full-btn");
          payFullBtn.onclick = () => {
            paymentAmountInput.value = installment.amountDue;
            updatePendingDisplay();
            paymentAmountInput.focus();
          };

          updatePendingDisplay();
          modal.classList.add("show");
        } else if (button.classList.contains("delete-activity-btn")) {
          const activityId = button.dataset.id;
          showConfirmation(
            "Delete Activity?",
            "Are you sure you want to delete this activity log?",
            async () => {
              try {
                await db.collection("activities").doc(activityId).delete();
                getEl(`activity-${activityId}`).remove();
                showToast("success", "Deleted", "Activity log removed.");
              } catch (error) {
                showToast("error", "Delete Failed", error.message);
              }
            }
          );
        } else if (button.classList.contains("delete-customer-btn")) {
          const customerId = button.dataset.id;
          showConfirmation(
            "Delete Customer?",
            "This will permanently delete this customer and all their loan data. This action cannot be undone.",
            async () => {
              try {
                await db.collection("customers").doc(customerId).delete();
                showToast(
                  "success",
                  "Customer Deleted",
                  "The customer record has been permanently removed."
                );
                await loadAndRenderAll();
              } catch (e) {
                showToast("error", "Delete Failed", e.message);
              }
            }
          );
        } else if (button.id === "delete-all-settled-btn") {
          showConfirmation(
            "Delete All Settled Accounts?",
            "This will permanently delete ALL settled and refinanced accounts from your records. This action cannot be undone. Are you sure?",
            async () => {
              const btn = getEl("delete-all-settled-btn");
              toggleButtonLoading(btn, true, "Deleting...");
              try {
                const querySnapshot = await db
                  .collection("customers")
                  .where("owner", "==", currentUser.uid)
                  .where("status", "in", ["settled", "Refinanced"])
                  .get();

                if (querySnapshot.empty) {
                  showToast(
                    "success",
                    "No Accounts",
                    "There are no settled accounts to delete."
                  );
                  toggleButtonLoading(btn, false);
                  return;
                }

                const batch = db.batch();
                querySnapshot.docs.forEach((doc) => {
                  batch.delete(doc.ref);
                });

                await batch.commit();
                showToast(
                  "success",
                  "Deletion Complete",
                  `${querySnapshot.size} settled account(s) have been deleted.`
                );
                await loadAndRenderAll();
              } catch (error) {
                showToast("error", "Deletion Failed", error.message);
              } finally {
                toggleButtonLoading(btn, false);
              }
            }
          );
        } else if (button.id === "clear-all-activities-btn") {
          showConfirmation(
            "Clear All Activities?",
            "This will delete all activity logs permanently. Are you sure?",
            async () => {
              const snapshot = await db
                .collection("activities")
                .where("owner_uid", "==", currentUser.uid)
                .get();
              const batch = db.batch();
              snapshot.docs.forEach((doc) => batch.delete(doc.ref));
              await batch.commit();
              recentActivities = [];
              renderActivityLog();
              showToast(
                "success",
                "All Cleared",
                "All activity logs have been deleted."
              );
            }
          );
        } else if (
          button.id === "logout-btn" ||
          button.id === "logout-settings-btn"
        ) {
          auth.signOut();
        } else if (button.id === "theme-toggle-btn") {
          if (window.toggleDarkMode) window.toggleDarkMode();
        } else if (button.id === "main-add-customer-btn") {
          getEl("customer-form").reset();
          getEl("customer-id").value = "";
          getEl("customer-form-modal-title").textContent = "Add New Customer";

          setAutomaticFirstDate();

          getEl("personal-info-fields").style.display = "block";
          getEl("kyc-info-fields").style.display = "block";
          getEl("loan-details-fields").style.display = "block";
          getEl("installment-preview").classList.add("hidden");

          document
            .querySelectorAll(".file-input-label span")
            .forEach((span) => {
              span.textContent = "Choose a file...";
            });
          getEl("customer-form-modal").classList.add("show");
        } else if (button.id === "edit-customer-info-btn") {
          const customer = [
            ...window.allCustomers.active,
            ...window.allCustomers.settled,
          ].find((c) => c.id === button.dataset.id);
          if (!customer) return;
          getEl("customer-form").reset();
          getEl("customer-id").value = customer.id;
          getEl("customer-name").value = customer.name;
          getEl("customer-phone").value = customer.phone || "";
          getEl("customer-father-name").value = customer.fatherName || "";
          getEl("customer-whatsapp").value = customer.whatsapp || "";
          getEl("customer-address").value = customer.address || "";

          getEl("customer-aadhar-number").value = customer.aadharNumber || "";
          getEl("customer-pan-number").value = customer.panNumber || "";

          getEl("personal-info-fields").style.display = "block";
          getEl("kyc-info-fields").style.display = "block";
          getEl("loan-details-fields").style.display = "none";
          getEl("installment-preview").classList.add("hidden");

          getEl("customer-form-modal-title").textContent = "Edit Customer Info";
          getEl("customer-details-modal").classList.remove("show");
          getEl("customer-form-modal").classList.add("show");
        } else if (button.id === "settle-loan-btn") {
          const currentLoanId = button.dataset.id;
          const currentCustomer = window.allCustomers.active.find(
            (c) => c.id === currentLoanId
          );
          if (!currentCustomer) return;

          const allActiveLoansForCustomer = window.allCustomers.active.filter(
            (c) => c.name === currentCustomer.name
          );

          if (allActiveLoansForCustomer.length <= 1) {
            showConfirmation(
              `Settle Loan?`,
              `This will move Finance ${
                currentCustomer.financeCount || 1
              } to the 'Settled' list. Are you sure?`,
              () => {
                settleLoanById(currentLoanId);
              }
            );
          } else {
            const optionsContainer = getEl("settle-options-container");
            optionsContainer.innerHTML = allActiveLoansForCustomer
              .map(
                (loan) => `
                    <div class="selection-item">
                        <input type="radio" name="settle-loan" id="settle-${
                          loan.id
                        }" value="${loan.id}" ${
                  loan.id === currentLoanId ? "checked" : ""
                }>
                        <label for="settle-${loan.id}">
                            Finance ${loan.financeCount || 1}
                            <span>Principal: ${formatCurrency(
                              loan.loanDetails.principal
                            )} / Due: ${formatCurrency(
                  loan.paymentSchedule[0]?.amountDue
                )}</span>
                        </label>
                    </div>
                `
              )
              .join("");

            getEl("settle-selection-modal").classList.add("show");
          }
        } else if (button.id === "settle-confirm-btn") {
          const selectedRadio = document.querySelector(
            'input[name="settle-loan"]:checked'
          );
          if (selectedRadio) {
            const loanIdToSettle = selectedRadio.value;
            settleLoanById(loanIdToSettle);
          } else {
            showToast(
              "error",
              "No Selection",
              "Please select a loan to settle."
            );
          }
        } else if (button.id === "add-new-loan-btn") {
          const customerId = button.dataset.id;
          const customer = window.allCustomers.active.find(
            (c) => c.id === customerId
          );
          if (!customer) return;

          getEl("new-loan-customer-id").value = customerId;
          getEl("new-loan-form").reset();

          // NEW: Call the automatic date function when opening the new loan modal
          setAutomaticNewLoanFirstDate();

          getEl("new-loan-installment-preview").classList.add("hidden");
          getEl("new-loan-modal").classList.add("show");
        } else if (button.id === "export-active-btn") {
          exportToExcel(
            window.allCustomers.active,
            "Active_Customers_Report.xlsx"
          );
        } else if (button.id === "export-settled-btn") {
          exportToExcel(
            window.allCustomers.settled,
            "Settled_Customers_Report.xlsx"
          );
        } else if (button.id === "export-backup-btn") {
          try {
            const backupData = {
              version: "2.0.0",
              exportedAt: new Date().toISOString(),
              customers: window.allCustomers,
            };
            const dataStr = JSON.stringify(backupData, null, 2);
            const blob = new Blob([dataStr], { type: "application/json" });
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = `loan-manager-backup-${
              new Date().toISOString().split("T")[0]
            }.json`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            showToast(
              "success",
              "Export Successful",
              "Your data has been downloaded."
            );
          } catch (e) {
            showToast("error", "Export Failed", e.message);
          }
        } else if (button.id === "import-backup-btn") {
          const fileInput = getEl("import-backup-input");
          const file = fileInput.files[0];
          if (!file) {
            showToast(
              "error",
              "No File",
              "Please choose a backup file to import."
            );
            return;
          }
          showConfirmation(
            "Overwrite All Data?",
            "Importing this file will permanently delete ALL your current customers and activities, replacing them with the data from the backup. This cannot be undone. Are you sure you want to proceed?",
            () => {
              const reader = new FileReader();
              reader.onload = async (event) => {
                try {
                  const backup = JSON.parse(event.target.result);
                  if (
                    !backup.customers ||
                    !backup.customers.active ||
                    !backup.customers.settled
                  ) {
                    throw new Error("Invalid backup file format.");
                  }
                  toggleButtonLoading(button, true, "Importing...");
                  const batch = db.batch();
                  const existingCustomers = await db
                    .collection("customers")
                    .where("owner", "==", currentUser.uid)
                    .get();
                  existingCustomers.docs.forEach((doc) =>
                    batch.delete(doc.ref)
                  );
                  [
                    ...backup.customers.active,
                    ...backup.customers.settled,
                  ].forEach((cust) => {
                    const newDocRef = db.collection("customers").doc();
                    const newCustData = {
                      ...cust,
                      owner: currentUser.uid,
                      createdAt: new Date(),
                    };
                    delete newCustData.id;
                    batch.set(newDocRef, newCustData);
                  });
                  await batch.commit();
                  showToast(
                    "success",
                    "Import Complete",
                    "Your data has been restored. Refreshing..."
                  );
                  await loadAndRenderAll();
                } catch (e) {
                  showToast("error", "Import Failed", e.message);
                } finally {
                  toggleButtonLoading(button, false);
                }
              };
              reader.readAsText(file);
            }
          );
        }
      }
    });

    document.body.addEventListener("submit", async (e) => {
      e.preventDefault();
      const form = e.target;
      if (form.id === "customer-form") {
        const id = getEl("customer-id").value;
        const saveBtn = getEl("customer-modal-save");
        toggleButtonLoading(saveBtn, true, id ? "Updating..." : "Saving...");

        try {
          const customerData = {
            name: getEl("customer-name").value,
            phone: getEl("customer-phone").value,
            fatherName: getEl("customer-father-name").value,
            address: getEl("customer-address").value,
            whatsapp: getEl("customer-whatsapp").value,
            aadharNumber: getEl("customer-aadhar-number").value.trim(),
            panNumber: getEl("customer-pan-number").value.trim().toUpperCase(),
          };

          const uploadFile = async (fileInputId, fileType) => {
            const file = getEl(fileInputId).files[0];
            if (!file) return null;
            const customerIdForPath = id || db.collection("customers").doc().id;
            const filePath = `kyc/${currentUser.uid}/${customerIdForPath}/${fileType}-${file.name}`;
            const snapshot = await storage.ref(filePath).put(file);
            return await snapshot.ref.getDownloadURL();
          };

          toggleButtonLoading(saveBtn, true, "Uploading Files...");
          const [aadharUrl, panUrl, picUrl, bankDetailsUrl] = await Promise.all(
            [
              uploadFile("customer-aadhar-file", "aadhar"),
              uploadFile("customer-pan-file", "pan"),
              uploadFile("customer-pic-file", "picture"),
              uploadFile("customer-bank-file", "bank"),
            ]
          );

          customerData.kycDocs = {
            ...(aadharUrl && { aadharUrl }),
            ...(panUrl && { panUrl }),
            ...(picUrl && { picUrl }),
            ...(bankDetailsUrl && { bankDetailsUrl }),
          };

          if (id) {
            const updatePayload = { ...customerData };
            Object.keys(updatePayload.kycDocs).forEach((key) => {
              if (updatePayload.kycDocs[key] === null) {
                delete updatePayload.kycDocs[key];
              }
            });
            const finalUpdate = {
              name: updatePayload.name,
              phone: updatePayload.phone,
              fatherName: updatePayload.fatherName,
              address: updatePayload.address,
              whatsapp: updatePayload.whatsapp,
              aadharNumber: getEl("customer-aadhar-number").value.trim(),
              panNumber: getEl("customer-pan-number")
                .value.trim()
                .toUpperCase(),
            };
            if (aadharUrl) finalUpdate["kycDocs.aadharUrl"] = aadharUrl;
            if (panUrl) finalUpdate["kycDocs.panUrl"] = panUrl;
            if (picUrl) finalUpdate["kycDocs.picUrl"] = picUrl;
            if (bankDetailsUrl)
              finalUpdate["kycDocs.bankDetailsUrl"] = bankDetailsUrl;

            await db.collection("customers").doc(id).update(finalUpdate);
            showToast(
              "success",
              "Customer Updated",
              "Details saved successfully."
            );
          } else {
            const customerName = getEl("customer-name").value.trim();
            const existingCustomer = [
              ...window.allCustomers.active,
              ...window.allCustomers.settled,
            ].find(
              (c) => c.name.trim().toLowerCase() === customerName.toLowerCase()
            );

            if (existingCustomer) {
              alert(
                "Customer already exists by this name. Please use another name or go to active accounts to add a new loan to that previous customer."
              );
              toggleButtonLoading(saveBtn, false);
              return;
            }

            const p = parseFloat(getEl("principal-amount").value);
            const r = parseFloat(getEl("interest-rate-modal").value);
            const freq = getEl("collection-frequency").value;
            const firstDate = getEl("first-collection-date").value;
            const endDate = getEl("loan-end-date").value;

            const n = calculateInstallments(firstDate, endDate, freq);

            if (isNaN(p) || isNaN(r) || isNaN(n) || !firstDate)
              throw new Error("Please fill all loan detail fields correctly.");

            customerData.loanDetails = {
              principal: p,
              interestRate: r,
              installments: n,
              frequency: freq,
              loanDate: new Date().toISOString().split("T")[0],
              firstCollectionDate: firstDate,
              loanEndDate: endDate,
              type: "simple_interest",
            };
            customerData.paymentSchedule = generateSimpleInterestSchedule(
              p,
              r,
              n,
              firstDate,
              endDate,
              freq
            );
            customerData.owner = currentUser.uid;
            customerData.createdAt =
              firebase.firestore.FieldValue.serverTimestamp();
            customerData.status = "active";
            customerData.financeCount = 1;

            toggleButtonLoading(saveBtn, true, "Saving Customer...");
            await db.collection("customers").add(customerData);
            await logActivity("NEW_LOAN", {
              customerName: customerData.name,
              amount: p,
            });
            showToast("success", "Customer Added", "New loan account created.");
          }
          getEl("customer-form-modal").classList.remove("show");
          await loadAndRenderAll();
        } catch (error) {
          console.error("Save/Update failed:", error);
          showToast("error", "Save Failed", error.message);
        } finally {
          toggleButtonLoading(saveBtn, false);
        }
      } else if (form.id === "payment-form") {
        const btn = getEl("payment-save-btn");
        toggleButtonLoading(btn, true, "Saving...");
        try {
          const customerId = getEl("payment-customer-id").value;
          const installmentNum = parseInt(
            getEl("payment-installment-number").value,
            10
          );
          const amountPaid = parseFloat(getEl("payment-amount").value);

          const customer = window.allCustomers.active.find(
            (c) => c.id === customerId
          );
          if (!customer) throw new Error("Customer not found");

          const updatedSchedule = JSON.parse(
            JSON.stringify(customer.paymentSchedule)
          );
          const instIndex = updatedSchedule.findIndex(
            (p) => p.installment === installmentNum
          );

          const installment = updatedSchedule[instIndex];
          const pending = installment.amountDue - amountPaid;
          installment.amountPaid = amountPaid;
          installment.pendingAmount = pending;
          installment.paidDate = new Date().toISOString();

          if (pending <= 0.001) {
            installment.status = "Paid";
            installment.pendingAmount = 0;
          } else if (amountPaid > 0) {
            installment.status = "Pending";
          } else {
            installment.status = "Due";
            installment.paidDate = null;
          }

          await db
            .collection("customers")
            .doc(customerId)
            .update({ paymentSchedule: updatedSchedule });
          await logActivity("PAYMENT_RECEIVED", {
            customerName: customer.name,
            amount: amountPaid,
          });

          showToast(
            "success",
            "Payment Saved",
            `Payment for installment #${installmentNum} recorded.`
          );
          getEl("payment-modal").classList.remove("show");
          await loadAndRenderAll();
          showCustomerDetails(customerId);
        } catch (e) {
          showToast("error", "Save Failed", e.message);
        } finally {
          toggleButtonLoading(btn, false);
        }
      } else if (form.id === "emi-calculator-form") {
        const p = parseFloat(getEl("calc-principal").value);
        const r = parseFloat(getEl("calc-rate").value);
        const n = parseInt(getEl("calc-tenure").value, 10);
        const freq = getEl("collection-frequency-calc").value;

        if (isNaN(p) || isNaN(r) || isNaN(n) || n <= 0) {
          showToast("error", "Invalid Input", "Please enter valid numbers.");
          return;
        }

        const tempInterest = (principal, rate, installments, frequency) => {
          const monthlyRateDecimal = rate / 100;
          if (frequency === "monthly")
            return principal * monthlyRateDecimal * installments;
          const dailyRate = monthlyRateDecimal / 30;
          if (frequency === "weekly")
            return principal * (dailyRate * 7) * installments;
          if (frequency === "daily")
            return principal * dailyRate * installments;
          return 0;
        };
        const totalInterest = tempInterest(p, r, n, freq);
        const totalPayment = p + totalInterest;
        const perInstallment = totalPayment / n;

        getEl("result-emi").textContent = formatCurrency(perInstallment);
        getEl("result-interest").textContent = formatCurrency(totalInterest);
        getEl("result-total").textContent = formatCurrency(totalPayment);
        getEl("calculator-results").classList.remove("hidden");
      } else if (form.id === "new-loan-form") {
        const saveBtn = getEl("new-loan-modal-save");
        toggleButtonLoading(saveBtn, true, "Creating...");
        try {
          const baseCustomerId = getEl("new-loan-customer-id").value;
          const allCustomerLoans = [
            ...window.allCustomers.active,
            ...window.allCustomers.settled,
          ];
          const baseCustomer = allCustomerLoans.find(
            (c) => c.id === baseCustomerId
          );
          if (!baseCustomer) throw new Error("Base customer data not found.");

          const allLoansForThisCustomerName = allCustomerLoans.filter(
            (c) => c.name === baseCustomer.name
          );
          const maxFinanceCount = Math.max(
            0,
            ...allLoansForThisCustomerName.map((c) => c.financeCount || 1)
          );

          const p = parseFloat(getEl("new-loan-principal").value);
          const r = parseFloat(getEl("new-loan-interest-rate").value);
          const freq = getEl("new-loan-frequency").value;
          const firstDate = getEl("new-loan-start-date").value;
          const endDate = getEl("new-loan-end-date").value;

          const n = calculateInstallments(firstDate, endDate, freq);

          if (isNaN(p) || isNaN(r) || isNaN(n) || !firstDate || !endDate)
            throw new Error("Please fill all new loan fields correctly.");

          const newLoanData = {
            name: baseCustomer.name,
            phone: baseCustomer.phone,
            fatherName: baseCustomer.fatherName,
            address: baseCustomer.address,
            whatsapp: baseCustomer.whatsapp,
            kycDocs: baseCustomer.kycDocs || {},
            owner: currentUser.uid,
            createdAt: firebase.firestore.FieldValue.serverTimestamp(),
            status: "active",
            financeCount: maxFinanceCount + 1,
            loanDetails: {
              principal: p,
              interestRate: r,
              installments: n,
              frequency: freq,
              loanDate: firstDate,
              firstCollectionDate: firstDate,
              loanEndDate: endDate,
              type: "simple_interest",
            },
            paymentSchedule: generateSimpleInterestSchedule(
              p,
              r,
              n,
              firstDate,
              endDate,
              freq
            ),
          };

          await db.collection("customers").add(newLoanData);
          await logActivity("NEW_LOAN", {
            customerName: newLoanData.name,
            amount: p,
            financeCount: newLoanData.financeCount,
          });
          showToast(
            "success",
            "Loan Added",
            `New Finance #${newLoanData.financeCount} created for ${newLoanData.name}.`
          );

          getEl("new-loan-modal").classList.remove("show");
          getEl("customer-details-modal").classList.remove("show");
          await loadAndRenderAll();
        } catch (error) {
          showToast("error", "Failed", error.message);
        } finally {
          toggleButtonLoading(saveBtn, false);
        }
      } else if (form.id === "change-password-form") {
        const btn = getEl("change-password-btn");
        toggleButtonLoading(btn, true, "Updating...");
        const currentPass = getEl("current-password").value;
        const newPass = getEl("new-password").value;
        const confirmPass = getEl("confirm-password").value;
        if (newPass !== confirmPass) {
          showToast("error", "Mismatch", "New passwords do not match.");
          toggleButtonLoading(btn, false);
          return;
        }
        if (newPass.length < 6) {
          showToast(
            "error",
            "Too Weak",
            "Password should be at least 6 characters long."
          );
          toggleButtonLoading(btn, false);
          return;
        }
        try {
          const user = auth.currentUser;
          const credential = firebase.auth.EmailAuthProvider.credential(
            user.email,
            currentPass
          );
          await user.reauthenticateWithCredential(credential);
          await user.updatePassword(newPass);
          showToast("success", "Success", "Password updated successfully.");
          form.reset();
        } catch (error) {
          showToast(
            "error",
            "Authentication Failed",
            "Incorrect current password or other error."
          );
        } finally {
          toggleButtonLoading(btn, false);
        }
      }
    });

    getEl("collection-frequency").addEventListener(
      "change",
      setAutomaticFirstDate
    );
    // NEW: Add listener for the new loan frequency dropdown
    getEl("new-loan-frequency").addEventListener(
      "change",
      setAutomaticNewLoanFirstDate
    );

    const loanDetailFields = [
      "principal-amount",
      "interest-rate-modal",
      "collection-frequency",
      "first-collection-date",
      "loan-end-date",
    ];
    loanDetailFields.forEach((id) => {
      const element = getEl(id);
      if (element) {
        element.addEventListener("input", updateInstallmentPreview);
        element.addEventListener("change", updateInstallmentPreview);
      }
    });

    const newLoanDetailFields = [
      "new-loan-principal",
      "new-loan-interest-rate",
      "new-loan-frequency",
      "new-loan-start-date",
      "new-loan-end-date",
    ];
    newLoanDetailFields.forEach((id) => {
      const element = getEl(id);
      if (element) {
        element.addEventListener("input", updateNewLoanInstallmentPreview);
        element.addEventListener("change", updateNewLoanInstallmentPreview);
      }
    });

    document.body.addEventListener("change", (e) => {
      if (e.target.id === "dark-mode-toggle") {
        if (window.toggleDarkMode) window.toggleDarkMode();
      } else if (e.target.id === "import-backup-input") {
        const fileName = e.target.files[0]
          ? e.target.files[0].name
          : "No file chosen";
        getEl("file-name-display").textContent = fileName;
      } else if (e.target.classList.contains("file-input")) {
        const label = e.target.nextElementSibling;
        const labelSpan = label.querySelector("span");
        const fileName =
          e.target.files.length > 0
            ? e.target.files[0].name
            : "Choose a file...";
        if (labelSpan) {
          labelSpan.textContent = fileName;
        }
      }
    });

    document.body.addEventListener("input", (e) => {
      if (e.target.id === "search-customers") {
        const term = e.target.value.toLowerCase();
        const filtered = window.allCustomers.active.filter((c) =>
          c.name.toLowerCase().includes(term)
        );
        renderActiveCustomerList(filtered);
      }
    });
  }

  initializeEventListeners();
});
