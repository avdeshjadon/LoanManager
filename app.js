// --- app.js ---
// Main application logic, UI rendering, aur data management

// Helper function to get element by ID
const getEl = (id) => document.getElementById(id);

// Helper function to show toast notifications
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

// Helper function to toggle button loading state
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
    // <-- Yeh line (if (btnText)) zaroori hai
    btnText.textContent = isLoading ? text : btnText.dataset.originalText;
  }
};

// --- Datepicker Logic (REMOVED) ---

// --- File Size Alert ---
const showSizeAlert = () => {
  if (document.querySelector(".popup-overlay")) return;

  const alertPopup = document.createElement("div");
  alertPopup.className = "popup-overlay";
  alertPopup.style.zIndex = "9999";
  alertPopup.innerHTML = `
      <div class="popup-content">
          <h2 style="color: var(--danger);">⚠️ File Too Large</h2>
          <p>The selected file exceeds the 1 MB size limit. Please choose a smaller file.</p>
          <button class="btn btn-primary" style="margin-top: 20px;" id="close-alert-btn">OK</button>
      </div>
  `;
  document.body.appendChild(alertPopup);

  const closePopup = () => alertPopup.remove();

  document
    .getElementById("close-alert-btn")
    .addEventListener("click", closePopup);
  alertPopup.addEventListener("click", (e) => {
    if (e.target === alertPopup) {
      closePopup();
    }
  });
};

// --- Confirmation Modal ---
const showConfirmation = (title, message, onConfirm) => {
  const modal = getEl("confirmation-modal");
  getEl("confirmation-title").textContent = title;
  getEl("confirmation-message").textContent = message;
  const prevZ = modal.style.zIndex;
  modal.style.zIndex = "10001";
  modal.classList.add("show");
  const confirmBtn = getEl("confirmation-confirm-btn");
  const cancelBtn = getEl("confirmation-cancel-btn");
  const cleanup = () => {
    modal.classList.remove("show");
    modal.style.zIndex = prevZ || "";
    confirmBtn.removeEventListener("click", confirmHandler);
    cancelBtn.removeEventListener("click", cancelHandler);
  };
  const confirmHandler = () => {
    try {
      if (typeof onConfirm === "function") onConfirm();
    } finally {
      cleanup();
    }
  };
  const cancelHandler = () => cleanup();
  confirmBtn.addEventListener("click", confirmHandler, { once: true });
  cancelBtn.addEventListener("click", cancelHandler, { once: true });
};

// --- Activity Logging ---
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

// --- Data Deletion / Cascade ---
const deleteStorageFolder = async (path) => {
  try {
    const folderRef = storage.ref(path);
    const list = await folderRef.listAll();
    await Promise.all(
      list.items.map(async (itemRef) => {
        try {
          await itemRef.delete();
        } catch (e) {
          console.warn(
            "Failed to delete storage item:",
            itemRef.fullPath,
            e.message || e
          );
        }
      })
    );
    for (const prefix of list.prefixes) {
      await deleteStorageFolder(prefix.fullPath);
    }
  } catch (err) {
    console.warn("Storage folder cleanup warning:", path, err.message || err);
  }
};

const deleteActivitiesByCustomerName = async (customerName) => {
  if (!currentUser || !customerName) return 0;
  try {
    const snap = await db
      .collection("activities")
      .where("owner_uid", "==", currentUser.uid)
      .where("details.customerName", "==", customerName)
      .get();
    if (snap.empty) return 0;
    const batch = db.batch();
    snap.docs.forEach((d) => batch.delete(d.ref));
    await batch.commit();
    return snap.size;
  } catch (e) {
    console.warn("Activity cleanup warning for", customerName, e.message || e);
    return 0;
  }
};

const deleteSingleCustomerCascade = async (customerId, customerData) => {
  try {
    let data = customerData;
    if (!data) {
      const doc = await db.collection("customers").doc(customerId).get();
      data = doc.exists ? doc.data() : null;
    }
    const customerName = data?.name;
    await deleteStorageFolder(`kyc/${currentUser.uid}/${customerId}`);
    await db.collection("customers").doc(customerId).delete();
    if (customerName) await deleteActivitiesByCustomerName(customerName);
    return { name: customerName };
  } catch (e) {
    throw e;
  }
};

const deleteAllCustomerRecordsByName = async (customerName) => {
  if (!currentUser || !customerName) return { count: 0 };
  const snap = await db
    .collection("customers")
    .where("owner", "==", currentUser.uid)
    .where("name", "==", customerName)
    .get();
  if (snap.empty) return { count: 0 };
  let deleted = 0;
  for (const doc of snap.docs) {
    try {
      await deleteStorageFolder(`kyc/${currentUser.uid}/${doc.id}`);
    } catch (_) {}
    try {
      await doc.ref.delete();
      deleted++;
    } catch (e) {
      console.warn("Failed to delete customer doc:", doc.id, e.message || e);
    }
  }
  await deleteActivitiesByCustomerName(customerName);
  return { count: deleted };
};

const renumberActiveLoansByCustomerName = async (customerName) => {
  if (!currentUser || !customerName) return 0;
  const snap = await db
    .collection("customers")
    .where("owner", "==", currentUser.uid)
    .where("name", "==", customerName)
    .where("status", "==", "active")
    .orderBy("createdAt", "asc")
    .get();
  if (snap.empty) return 0;
  const batch = db.batch();
  snap.docs.forEach((doc, idx) => {
    batch.update(doc.ref, { financeCount: idx + 1 });
  });
  await batch.commit();
  return snap.size;
};

// --- Stats & Sidebar ---
const calculateKeyStats = (activeLoans, settledLoans) => {
  let totalPrincipal = 0,
    totalOutstanding = 0,
    totalInterestEarned = 0;

  [...activeLoans, ...settledLoans].forEach((c) => {
    if (c.loanDetails && c.paymentSchedule) {
      totalPrincipal += c.loanDetails.principal;
      const endBoundDate = parseDateFlexible(c.loanDetails.loanEndDate);
      const totalInterest = calculateTotalInterest(
        c.loanDetails.principal,
        c.loanDetails.interestRate,
        c.loanDetails.loanGivenDate,
        endBoundDate
      );
      const totalRepayable = c.loanDetails.principal + totalInterest;
      const totalPaid = c.paymentSchedule.reduce(
        (sum, p) => sum + p.amountPaid,
        0
      );
      totalInterestEarned += Math.max(0, totalInterest);

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

// --- Main Data Load & Render ---
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

// --- UI Rendering Functions ---

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
        case "LOAN_RESTORED":
          icon = "fa-undo-alt text-warning";
          text = `Loan restored for ${customerName}`;
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

const formatForDisplay = (val) => {
  if (!val) return "N/A";
  try {
    const d = parseDateFlexible(val);
    if (!d || isNaN(+d)) return "N/A";
    return formatForInput({ id: "any" }, d);
  } catch (_) {
    return "N/A";
  }
};

const populateTodaysCollection = () => {
  const container = getEl("todays-collection-container");
  if (!container) return;

  let todaysEmis = [];
  const today = new Date();
  const todayStr = today.toISOString().split("T")[0];

  window.allCustomers.active
    .filter((c) => c.loanDetails && c.paymentSchedule)
    .forEach((customer) => {
      customer.paymentSchedule.forEach((inst) => {
        if (
          (inst.status === "Due" || inst.status === "Pending") &&
          inst.dueDate === todayStr
        ) {
          todaysEmis.push({
            name: `${customer.name} (F${customer.financeCount || 1})`,
            amount: inst.pendingAmount || inst.amountDue || 0,
            date: inst.dueDate,
            customerId: customer.id,
            isMonthly: customer.loanTermType === "monthly",
          });
        }
      });
    });

  todaysEmis.sort((a, b) => a.name.localeCompare(b.name));

  // Calculate total amount due for today and count
  const totalDueToday = todaysEmis.reduce(
    (sum, it) => sum + (Number(it.amount) || 0),
    0
  );
  const totalCount = todaysEmis.length;

  // Build total summary box HTML
  const totalBoxHtml = `
    <div class="today-total-box" role="region" aria-label="Today's total collection">
      <div class="today-total-left">
        <div class="today-total-label">Today's Total Collection</div>
        <div class="today-total-sub">Collections: ${totalCount}</div>
      </div>
      <div class="today-total-amount">${formatCurrency(totalDueToday)}</div>
    </div>
  `;

  if (todaysEmis.length === 0) {
    container.innerHTML = `
      ${totalBoxHtml}
      <ul class="activity-list">
        <li class="activity-item" style="cursor:default; justify-content:center;">No collections due today.</li>
      </ul>`;
    return;
  }

  container.innerHTML = `
    ${totalBoxHtml}
    <ul class="activity-list">
      ${todaysEmis
        .map((item) => {
          const formattedDate = formatForDisplay(item.date);
          const monthlyBadge = item.isMonthly
            ? '<span class="finance-status-badge monthly-loan" style="font-size: 0.7rem; margin-left: 8px;">Monthly</span>'
            : "";
          return `<li class="activity-item" data-id="${
            item.customerId
          }"><div class="activity-info"><span class="activity-name">${
            item.name
          }${monthlyBadge}</span><span class="activity-date">${formattedDate}</span></div><div class="activity-value"><span class="activity-amount">${formatCurrency(
            item.amount
          )}</span></div></li>`;
        })
        .join("")}
    </ul>`;
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
    "todays-collection-section"
  ).innerHTML = `<div class="form-card"><h3><i class="fas fa-calendar-day"></i> Today's Collection Summary</h3><div id="todays-collection-container" class="activity-container"></div></div>`;

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

    // NEW: Add Monthly Loan badge
    const isMonthly = c.loanTermType === "monthly";
    const monthlyBadge = isMonthly
      ? '<span class="finance-status-badge monthly-loan">Monthly Loan</span>'
      : "";

    const financeStatusText = `Finance ${financeCount} - ${c.status}`;
    const nameHtml = `<div class="customer-name">${c.name} <span class="finance-status-badge">${financeStatusText}</span>${monthlyBadge}</div>`;
    const detailsHtml = `<span>Principal: ${formatCurrency(
      c.loanDetails.principal
    )}</span><span class="list-profit-display success">Interest: ${formatCurrency(
      interestEarned
    )}</span>`;

    const restoreButton = `<button class="btn btn-success btn-sm restore-customer-btn" data-id="${c.id}" title="Restore Loan"><i class="fas fa-undo-alt"></i></button>`;
    const deleteButton = `<button class="btn btn-danger btn-sm delete-customer-btn" data-id="${c.id}" title="Delete Customer"><i class="fas fa-trash-alt"></i></button>`;

    li.innerHTML = `<div class="customer-info" data-id="${c.id}">${nameHtml}<div class="customer-details">${detailsHtml}</div></div><div class="customer-actions">${restoreButton}${deleteButton}<span class="view-details-prompt" data-id="${c.id}">View Details</span></div>`;
    element.appendChild(li);
  });
};

const getSortValue = (loan, key) => {
  switch (key) {
    case "lastCollectionDate": {
      // Find the most recent date a payment was recorded
      const lastPaid = loan.paymentSchedule
        ?.filter((p) => p.status === "Paid" && p.paidDate)
        .sort((a, b) => new Date(b.paidDate) - new Date(a.paidDate))[0];
      // Use loan given date as a fallback if no payments exist
      const fallbackDate = loan.loanDetails?.loanGivenDate
        ? new Date(loan.loanDetails.loanGivenDate)
        : new Date(0);
      return lastPaid ? new Date(lastPaid.paidDate) : fallbackDate;
    }
    case "firstCollectionDate": {
      return loan.loanDetails?.firstCollectionDate
        ? new Date(loan.loanDetails.firstCollectionDate)
        : new Date(0);
    }
    case "name":
    default:
      return loan.name.toLowerCase();
  }
};

const renderActiveCustomerList = (customerArray, sortKey = "name") => {
  const listEl = getEl("customers-list");
  if (!listEl) return;

  const customerGroups = new Map();
  customerArray.forEach((customer) => {
    if (!customerGroups.has(customer.name)) {
      customerGroups.set(customer.name, []);
    }
    customerGroups.get(customer.name).push(customer);
  });

  // Convert Map values to an array for sorting
  const groupedCustomers = Array.from(customerGroups.entries()).map(
    ([name, loans]) => {
      // Sort loans within the group by financeCount to always show the latest loan's ID
      const latestLoan = loans.sort(
        (a, b) => (b.financeCount || 1) - (a.financeCount || 1)
      )[0];

      // Determine the sort value based on the chosen key, using the latest/representative loan
      const sortValue = getSortValue(latestLoan, sortKey);

      return { name, loans, latestLoan, sortValue };
    }
  );

  // Sort the grouped customers
  groupedCustomers.sort((a, b) => {
    if (sortKey === "name") {
      return a.sortValue.localeCompare(b.sortValue);
    } else {
      // Sort by date/time (most recent first)
      return b.sortValue.getTime() - a.sortValue.getTime();
    }
  });

  if (groupedCustomers.length === 0) {
    listEl.innerHTML = `<li class="activity-item" style="cursor:default; justify-content:center;"><p>No customers found.</p></li>`;
    return;
  }

  let listHtml = "";
  for (const { name, loans, latestLoan } of groupedCustomers) {
    const totalActiveLoans = loans.length;

    const totalOutstanding = loans.reduce((sum, loan) => {
      if (!loan.loanDetails || !loan.paymentSchedule) return sum;
      const totalInterest = calculateTotalInterest(
        loan.loanDetails.principal,
        loan.loanDetails.interestRate,
        loan.loanDetails.loanGivenDate,
        loan.loanDetails.loanEndDate
      );
      const totalRepayable = loan.loanDetails.principal + totalInterest;
      const totalPaid = loan.paymentSchedule.reduce(
        (s, p) => s + (p.amountPaid || 0),
        0
      );
      return sum + Math.max(0, totalRepayable - totalPaid);
    }, 0);

    const nameHtml = `<div class="customer-name">${name}</div>`;
    const detailsHtml = `<span>Total Outstanding: ${formatCurrency(
      totalOutstanding
    )}</span>`;
    const loanCountBadge =
      totalActiveLoans > 1
        ? `<span class="finance-count-badge">${totalActiveLoans}</span>`
        : "";

    // NEW: Add Monthly Loan badge if the LATEST loan is monthly
    const isMonthly = latestLoan.loanTermType === "monthly";
    const monthlyBadge = isMonthly
      ? '<span class="finance-status-badge monthly-loan">Monthly Loan</span>'
      : "";

    listHtml += `
          <li class="customer-item">
              <div class="customer-info" data-id="${latestLoan.id}">
                  ${nameHtml}
                  <div class="customer-details">${detailsHtml}</div>
              </div>
              <div class="customer-actions">
                  ${monthlyBadge}
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
  ).innerHTML = `<div class="form-card"><div class="card-header"><h3>Active Accounts</h3><div class="form-row" style="grid-template-columns: 1fr 1fr; gap: 1rem; width: 100%; max-width: 400px; margin-left: auto;"><div class="form-group" style="margin-bottom: 0;"><select id="customer-sort-select" class="form-control"><option value="name">Sort by: Name</option><option value="first">Sort by: First Date</option><option value="last">Sort by: Last Date</option></select></div><button class="btn btn-outline" id="export-active-btn" style="height: 48px;"><i class="fas fa-file-excel"></i> Export</button></div></div><div class="form-group"><input type="text" id="search-customers" class="form-control" placeholder="Search active customers..." /></div><ul id="customers-list" class="customer-list"></ul></div>`;
  renderActiveCustomerList(window.allCustomers.active, activeSortKey);
  getEl("customer-sort-select").value = activeSortKey; // Set selected value

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
          isMonthly: customer.loanTermType === "monthly",
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
      .map((item) => {
        const formattedDate = formatForDisplay(item.date);
        // NEW: Add monthly badge
        const monthlyBadge = item.isMonthly
          ? '<span class="finance-status-badge monthly-loan" style="font-size: 0.7rem; margin-left: 8px;">Monthly</span>'
          : "";
        return `<li class="activity-item" data-id="${
          item.customerId
        }"><div class="activity-info"><span class="activity-name">${
          item.name
        }${monthlyBadge}</span><span class="activity-date">${formattedDate}</span></div><div class="activity-value"><span class="activity-amount">${formatCurrency(
          item.amount
        )}</span></div></li>`;
      })
      .join("")}</ul>`;
  };
  upcomingContainer.innerHTML = renderEmiList(upcoming);
  overdueContainer.innerHTML = renderEmiList(overdue);
};

const showCustomerDetails = (customerId) => {
  const customer = [
    ...window.allCustomers.active,
    ...window.allCustomers.settled,
  ].find((c) => c.id === customerId);
  if (!customer) return;

  const loanListSource =
    customer.status === "active"
      ? window.allCustomers.active
      : window.allCustomers.settled;

  const allLoansForCustomer = loanListSource
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
              Finance ${loan.financeCount || 1} ${
            loan.loanTermType === "monthly" ? "(Monthly)" : ""
          }
          </div>`
      )
      .join("");
    const currentLoan =
      loansToDisplay.find((l) => l.id === customerId) || customer;
    triggerText.textContent = `Finance ${currentLoan.financeCount || 1} ${
      currentLoan.loanTermType === "monthly" ? "(Monthly)" : ""
    }`;
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
  if (frequencyBadge) {
    if (customer.loanTermType === "monthly") {
      frequencyBadge.textContent = "Monthly Loan";
      frequencyBadge.className =
        "loan-frequency-badge frequency-monthly-special"; // You can style this
    } else if (customer.loanDetails?.frequency) {
      frequencyBadge.textContent = customer.loanDetails.frequency;
      frequencyBadge.className = `loan-frequency-badge frequency-${customer.loanDetails.frequency}`;
    } else {
      frequencyBadge.textContent = "N/A";
      frequencyBadge.className = "loan-frequency-badge";
    }
  }

  const mopBadge = getEl("details-modal-loan-mop");
  if (mopBadge && customer.loanDetails?.modeOfPayment) {
    mopBadge.textContent = customer.loanDetails.modeOfPayment;
  } else if (mopBadge) {
    mopBadge.textContent = "N/A";
  }

  getEl("generate-pdf-btn").dataset.id = customerId;
  getEl("send-whatsapp-btn").dataset.id = customerId;

  if (!customer.loanDetails || !customer.paymentSchedule) {
    modalBody.innerHTML = `<p style="padding: 2rem; text-align: center;">This customer has incomplete or outdated loan data.</p>`;
    return;
  }

  const { paymentSchedule: schedule, loanDetails: details } = customer;
  const dispDate = (val) => {
    if (!val) return "N/A";
    try {
      const d = parseDateFlexible(val);
      if (!d || isNaN(+d)) return "N/A";
      return formatForInput({ id: "any" }, d);
    } catch (_) {
      return "N/A";
    }
  };

  const totalInterest = calculateTotalInterest(
    details.principal,
    details.interestRate,
    details.loanGivenDate,
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

  const bankNameDisplay = customer.bankName || "N/A";
  const accountNumberDisplay = customer.accountNumber || "N/A";
  const ifscDisplay = customer.ifsc || "N/A";

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
  const aadharFrontButton = createKycViewButton(
    kycDocs.aadharUrlFront,
    "View Front"
  );
  const aadharBackButton = createKycViewButton(
    kycDocs.aadharUrlBack,
    "View Back"
  );
  const panButton = createKycViewButton(kycDocs.panUrl);
  const picButton = createKycViewButton(kycDocs.picUrl, "View Photo");
  const bankButton = createKycViewButton(kycDocs.bankDetailsUrl);

  // Calculate all loans total for the display
  const allLoans = [
    ...window.allCustomers.active,
    ...window.allCustomers.settled,
  ].filter(
    (c) => c.name === customer.name && c.loanDetails && c.paymentSchedule
  );
  let totalP = 0,
    totalI = 0,
    totalPaidAll = 0;
  allLoans.forEach((c) => {
    const li = c.loanDetails;
    const interest = calculateTotalInterest(
      li.principal,
      li.interestRate,
      li.loanGivenDate,
      li.loanEndDate
    );
    totalP += Number(li.principal || 0);
    totalI += Number(interest || 0);
    totalPaidAll += c.paymentSchedule.reduce(
      (s, p) => s + (p.amountPaid || 0),
      0
    );
  });
  const totalOutstandingAll = Math.max(0, totalP + totalI - totalPaidAll);

  let avatarHtml;
  if (kycDocs.picUrl) {
    avatarHtml = `<img src="${kycDocs.picUrl}" alt="${customer.name.charAt(
      0
    )}" class="profile-avatar-img">`;
  } else {
    avatarHtml = customer.name.charAt(0).toUpperCase();
  }

  modalBody.innerHTML = `<div class="details-view-grid"><div class="customer-profile-panel"><div class="profile-header">

  <div class="profile-avatar">${avatarHtml}</div>

  <h3 class="profile-name">${customer.name}</h3><p class="profile-contact">${
    customer.phone || "N/A"
  }</p></div><div class="profile-section"><h4>Personal & Loan Details</h4>
  <div class="profile-stat"><span class="label">Loan Given Date</span><span class="value">${dispDate(
    details.loanGivenDate
  )}</span></div>
  <div class="profile-stat"><span class="label">Date of Birth</span><span class="value">${dispDate(
    customer.dob
  )}</span></div>
  <div class="profile-stat"><span class="label">First Collection</span><span class="value">${dispDate(
    details.firstCollectionDate
  )}</span></div>
  <div class="profile-stat"><span class="label">Last Loan Date</span><span class="value">${dispDate(
    schedule && schedule.length > 0
      ? schedule[schedule.length - 1].dueDate
      : null
  )}</span></div>
  </div>
  <div class="profile-section"><h4>Customer Loan Totals</h4>
      <div class="profile-stat"><span class="label">Total Principal (All Loans)</span><span class="value">${formatCurrency(
        totalP
      )}</span></div>
      <div class="profile-stat"><span class="label">Total Interest (All Loans)</span><span class="value">${formatCurrency(
        totalI
      )}</span></div>
      <div class="profile-stat"><span class="label">Outstanding (All Loans)</span><span class="value">${formatCurrency(
        totalOutstandingAll
      )}</span></div>
  </div>
  <div class="profile-section"><h4>KYC Documents</h4>
  <div class="profile-stat"><span class="label">Aadhar Card (Front)</span>${aadharFrontButton}</div>
  <div class="profile-stat"><span class="label">Aadhar Card (Back)</span>${aadharBackButton}</div>
  <div class="profile-stat"><span class="label">PAN Card</span>${panButton}</div>
  <div class="profile-stat"><span class="label">Client Photo</span>${picButton}</div>
  <div class="profile-stat"><span class="label">Bank Details</span>${bankButton}</div>
  <div class="profile-stat"><span class="label">Bank Name</span><span class="value">${bankNameDisplay}</span></div>
  <div class="profile-stat"><span class="label">Account Number</span><span class="value">${accountNumberDisplay}</span></div>
  <div class="profile-stat"><span class="label">IFSC</span><span class="value">${ifscDisplay}</span></div>
  <div class="profile-stat"><span class="label">Father's Name</span><span class="value">${
    customer.fatherName || "N/A"
  }</span></div>
  <div class="profile-stat profile-stat-address"><span class="label">Address</span><span class="value address-value">${
    customer.address || "N/A"
  }</span></div>
  </div><div class="loan-progress-section"><h4>Loan Progress (${paidInstallments} of ${
    schedule.length
  } Paid)</h4><div class="progress-bar"><div class="progress-bar-inner" style="width: ${progress}%;"></div></div></div><div class="loan-actions"><button class="btn btn-outline" id="edit-customer-info-btn" data-id="${
    customer.id
  }"><i class="fas fa-edit"></i> Edit Info</button>${actionButtons}</div></div><div class="emi-schedule-panel"><div class="emi-table-container"><table class="emi-table"><thead><tr><th>#</th><th>Due Date</th><th>Amount Due</th><th>Amount Paid</th><th>MoP</th><th>Status</th><th class="no-pdf">Action</th></tr></thead><tbody id="emi-schedule-body-details"></tbody></table></div><div class="loan-summary-box"><h4>Loan Summary</h4><div class="calc-result-item"><span>Principal Amount</span><span>${formatCurrency(
    details.principal
  )}</span></div><div class="calc-result-item"><span>Interest Rate (Monthly)</span><span>${
    (details.interestRate ?? 0) + "%"
  }</span></div><div class="calc-result-item"><span>Outstanding Amount</span><span>${formatCurrency(
    remainingToCollect
  )}</span></div></div><div class="modal-summary-stats"><div class="summary-stat-item"><span class="label">Amount Received</span><span class="value received">${formatCurrency(
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
    if (customer.loanTermType === "monthly" && inst.status === "Pending") {
      statusText = `PRINCIPAL DUE (${formatCurrency(inst.pendingAmount)} due)`;
    } else if (inst.status === "Pending") {
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
    if (customer.status === "active" && inst.status === "Paid") {
      actionButtons += `<button class="btn btn-danger btn-sm undo-payment-btn" data-installment="${inst.installment}" data-id="${customer.id}" title="Undo Payment"><i class="fas fa-undo"></i></button>`;
    }

    const paymentMop = inst.amountPaid > 0 ? inst.modeOfPayment || "N/A" : "";
    const displayedDue = dispDate(inst.dueDate);

    tr.innerHTML = `<td>${inst.installment}</td>
      <td>${displayedDue}</td>
      <td>${formatCurrency(inst.amountDue)}</td>
      <td>${formatCurrency(inst.amountPaid)}</td>
      <td>${paymentMop}</td>
      <td><span class="emi-status status-${statusClass}">${statusText}</span></td>
      <td class="no-pdf">${actionButtons}</td>`;
    emiTableBody.appendChild(tr);
  });
};

const canEditLoanDetailsForCustomer = (customer) => {
  if (!customer || !customer.paymentSchedule) return true;
  // Check if any installment has been paid (amountPaid > 0)
  return !customer.paymentSchedule.some(
    (installment) => parseFloat(installment.amountPaid || 0) > 0
  );
};

// --- Form & Modal Logic ---

function setLoanDetailFieldsRequired(isRequired) {
  const ids = [
    "principal-amount",
    "interest-rate-modal",
    "collection-frequency",
    "first-collection-date",
    "loan-end-date",
    "loan-given-date",
    "loan-mop",
  ];
  ids.forEach((id) => {
    const el = document.getElementById(id);
    if (!el) return;
    if (isRequired) el.setAttribute("required", "required");
    else el.removeAttribute("required");
  });
}
window.setLoanDetailFieldsRequired = setLoanDetailFieldsRequired;

// Safe setters
function safeSetText(target, value) {
  const el =
    typeof target === "string" ? document.getElementById(target) : target;
  if (el) el.textContent = value;
}
function safeSetValue(id, value) {
  const el = document.getElementById(id);
  if (el) el.value = value;
}

async function settleLoanById(loanId) {
  const customerToSettle = window.allCustomers.active.find(
    (c) => c.id === loanId
  );
  if (!customerToSettle) {
    showToast("error", "Not Found", "The selected loan could not be found.");
    return;
  }
  try {
    await db.collection("customers").doc(loanId).update({ status: "settled" });
    await logActivity("LOAN_SETTLED", {
      customerName: customerToSettle.name,
      financeCount: customerToSettle.financeCount || 1,
    });
    try {
      await renumberActiveLoansByCustomerName(customerToSettle.name);
    } catch (e) {
      console.warn("Renumbering after settle failed:", e);
    }
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

async function restoreLoanById(loanId) {
  const customerToRestore = window.allCustomers.settled.find(
    (c) => c.id === loanId
  );
  if (!customerToRestore) {
    showToast("error", "Not Found", "The selected loan could not be found.");
    return;
  }
  try {
    await db.collection("customers").doc(loanId).update({ status: "active" });

    await logActivity("LOAN_RESTORED", {
      customerName: customerToRestore.name,
      financeCount: customerToRestore.financeCount || 1,
    });

    try {
      await renumberActiveLoansByCustomerName(customerToRestore.name);
    } catch (e) {
      console.warn("Renumbering after restore failed:", e);
    }

    showToast(
      "success",
      "Loan Restored",
      `Finance ${customerToRestore.financeCount || 1} moved back to active.`
    );

    await loadAndRenderAll();
  } catch (error) {
    showToast("error", "Restore Failed", error.message);
  }
}

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
    const loanGivenDate = getEl("loan-given-date")?.value || new Date();
    const totalInterest = calculateTotalInterest(p, r, loanGivenDate, endDate);
    const totalRepayable = p + totalInterest;
    const minInstallment = +(totalRepayable / n).toFixed(2);

    previewDiv.innerHTML = `
      <div class="calc-flex">
        <p style="margin:0 0 .5rem 0">Calculated: <strong>${n} installments</strong> of approx. <strong>${formatCurrency(
      minInstallment
    )}</strong> each.</p>
        <div class="form-row" style="align-items:end; grid-template-columns: 1fr 1fr; gap: 1rem;">
          <div class="form-group" style="margin:0">
            <label for="custom-installment-amount">Choose per-installment amount (min ${formatCurrency(
              minInstallment
            )})</label>
            <input type="number" step="0.01" min="${minInstallment}" value="${minInstallment}" id="custom-installment-amount" class="form-control" />
          </div>
          <div class="form-group" style="margin:0">
            <div id="custom-emi-summary" class="calc-summary"></div>
          </div>
        </div>
        <input type="hidden" id="chosen-n-installments" value="${n}" />
        <input type="hidden" id="chosen-end-date" value="${endDate}" />
      </div>`;

    const summaryEl = getEl("custom-emi-summary");
    const inputEl = getEl("custom-installment-amount");
    const hiddenN = getEl("chosen-n-installments");
    const hiddenEnd = getEl("chosen-end-date");

    const applyValue = () => {
      let raw = inputEl.value;
      if (raw === "" || raw === null) {
        summaryEl.innerHTML = `<span>Enter an amount ≥ ${formatCurrency(
          minInstallment
        )}</span>`;
        inputEl.classList.remove("invalid");
        return;
      }
      let chosen = parseFloat(raw);
      if (isNaN(chosen)) {
        inputEl.classList.add("invalid");
        return;
      }
      if (chosen < minInstallment) {
        inputEl.classList.add("invalid");
        summaryEl.innerHTML = `<span style="color:var(--danger)">Amount is below minimum ${formatCurrency(
          minInstallment
        )}</span>`;
        return;
      }
      inputEl.classList.remove("invalid");
      const totalCents = Math.round(totalRepayable * 100);
      const chosenCents = Math.round(chosen * 100);
      const newN = Math.max(1, Math.ceil(totalCents / chosenCents));
      const newEnd = computeEndDateFromInstallments(firstDate, newN, freq);
      hiddenN.value = String(newN);
      hiddenEnd.value = newEnd;
      const remainderCents = totalCents - chosenCents * (newN - 1);
      const lastAmt = +(remainderCents / 100).toFixed(2);
      const lastInfo =
        lastAmt !== chosen
          ? ` &nbsp;|&nbsp; <span>Last: <strong>${formatCurrency(
              lastAmt
            )}</strong></span>`
          : "";
      summaryEl.innerHTML = `<span><strong>${newN}</strong> installments</span> &nbsp;|&nbsp; <span>Per installment: <strong>${formatCurrency(
        chosen
      )}</strong></span>${lastInfo} &nbsp;|&nbsp; <span>Ends on: <strong>${newEnd}</strong></span>`;
    };
    inputEl.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter") {
        ev.preventDefault();
        applyValue();
      }
    });
    inputEl.addEventListener("blur", applyValue);
    inputEl.value = String(minInstallment);
    applyValue();
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
    const loanGivenDate =
      getEl("new-loan-given-date")?.value ||
      getEl("loan-given-date")?.value ||
      new Date();
    const totalInterest = calculateTotalInterest(p, r, loanGivenDate, endDate);
    const totalRepayable = p + totalInterest;
    const minInstallment = +(totalRepayable / n).toFixed(2);

    previewDiv.innerHTML = `
      <div class="calc-flex">
        <p style="margin:0 0 .5rem 0">New Loan: <strong>${n} installments</strong> of approx. <strong>${formatCurrency(
      minInstallment
    )}</strong> each.</p>
        <div class="form-row" style="align-items:end; grid-template-columns: 1fr 1fr; gap: 1rem;">
          <div class="form-group" style="margin:0">
            <label for="new-loan-custom-installment-amount">Choose per-installment amount (min ${formatCurrency(
              minInstallment
            )})</label>
            <input type="number" step="0.01" min="${minInstallment}" value="${minInstallment}" id="new-loan-custom-installment-amount" class="form-control" />
          </div>
          <div class="form-group" style="margin:0">
            <div id="new-loan-custom-emi-summary" class="calc-summary"></div>
          </div>
        </div>
        <input type="hidden" id="new-loan-chosen-n-installments" value="${n}" />
        <input type="hidden" id="new-loan-chosen-end-date" value="${endDate}" />
      </div>`;

    const summaryEl = getEl("new-loan-custom-emi-summary");
    const inputEl = getEl("new-loan-custom-installment-amount");
    const hiddenN = getEl("new-loan-chosen-n-installments");
    const hiddenEnd = getEl("new-loan-chosen-end-date");

    const applyValue = () => {
      let raw = inputEl.value;
      if (raw === "" || raw === null) {
        summaryEl.innerHTML = `<span>Enter an amount ≥ ${formatCurrency(
          minInstallment
        )}</span>`;
        inputEl.classList.remove("invalid");
        return;
      }
      let chosen = parseFloat(raw);
      if (isNaN(chosen)) {
        inputEl.classList.add("invalid");
        return;
      }
      if (chosen < minInstallment) {
        inputEl.classList.add("invalid");
        summaryEl.innerHTML = `<span style="color:var(--danger)"><strong>Amount is below minimum ${formatCurrency(
          minInstallment
        )}</strong></span>`;
        return;
      }
      inputEl.classList.remove("invalid");
      const totalCents = Math.round(totalRepayable * 100);
      const chosenCents = Math.round(chosen * 100);
      const newN = Math.max(1, Math.ceil(totalCents / chosenCents));
      const newEnd = computeEndDateFromInstallments(firstDate, newN, freq);
      hiddenN.value = String(newN);
      hiddenEnd.value = newEnd;
      const endDateField = getEl("new-loan-end-date");
      if (endDateField) endDateField.value = formatForDateInput(parseDateFlexible(newEnd)); // MODIFIED
      const remainderCents = totalCents - chosenCents * (newN - 1);
      const lastAmt = +(remainderCents / 100).toFixed(2);
      const lastInfo =
        lastAmt !== chosen
          ? ` &nbsp;|&nbsp; <span>Last: <strong>${formatCurrency(
              lastAmt
            )}</strong></span>`
          : "";
      summaryEl.innerHTML = `<span><strong>${newN}</strong> installments</span> &nbsp;|&nbsp; <span>Per installment: <strong>${formatCurrency(
        chosen
      )}</strong></span>${lastInfo} &nbsp;|&nbsp; <span>Ends on: <strong>${newEnd}</strong></span>`;
    };
    inputEl.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter") {
        ev.preventDefault();
        applyValue();
      }
    });
    inputEl.addEventListener("blur", applyValue);
    inputEl.value = String(minInstallment);
    applyValue();
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
  const givenInput = getEl("loan-given-date");
  const base =
    givenInput && givenInput.value
      ? parseDateFlexible(givenInput.value)
      : new Date();

  if (freq === "daily") {
    base.setDate(base.getDate() + 1);
  } else if (freq === "weekly") {
    base.setDate(base.getDate() + 7);
  } else if (freq === "monthly") {
    base.setMonth(base.getMonth() + 1);
  }
  firstDateInput.value = formatForDateInput(base); // MODIFIED
  firstDateInput.dispatchEvent(new Event("change"));
}

function setAutomaticNewLoanFirstDate() {
  const freq = getEl("new-loan-frequency").value;
  const firstDateInput = getEl("new-loan-start-date");
  const givenInput = getEl("new-loan-given-date");
  const base =
    givenInput && givenInput.value
      ? parseDateFlexible(givenInput.value)
      : new Date();

  if (freq === "daily") {
    base.setDate(base.getDate() + 1);
  } else if (freq === "weekly") {
    base.setDate(base.getDate() + 7);
  } else if (freq === "monthly") {
    base.setMonth(base.getMonth() + 1);
  }
  firstDateInput.value = formatForDateInput(base); // MODIFIED
  firstDateInput.dispatchEvent(new Event("change"));
  try {
    const endDefault = computeEndDateFromInstallments(
      firstDateInput.value,
      2,
      freq
    );
    const endInput = getEl("new-loan-end-date");
    if (endInput) endInput.value = formatForDateInput(parseDateFlexible(endDefault)); // MODIFIED
  } catch (_) {}
}

// --- Payment Logic ---

const undoInstallmentPayment = async (customerId, installmentNum) => {
  const docRef = db.collection("customers").doc(customerId);
  const snap = await docRef.get();
  if (!snap.exists) throw new Error("Customer not found.");
  const data = snap.data();
  const schedule = (data.paymentSchedule || []).map((inst) => ({ ...inst }));
  const idx = schedule.findIndex((p) => p.installment === installmentNum);
  if (idx === -1) throw new Error("Installment not found.");
  const inst = schedule[idx];
  if (inst.status !== "Paid")
    throw new Error("Only fully paid installments can be undone.");
  // Revert values (status back to 'Due')
  inst.amountPaid = 0;
  inst.pendingAmount = inst.amountDue;
  inst.status = "Due";
  inst.paidDate = null;
  inst.modeOfPayment = null;
  await docRef.update({ paymentSchedule: schedule });
  showToast(
    "success",
    "Reverted",
    `Installment #${installmentNum} reverted to Unpaid.`
  );
};

const recordPayment = async (
  customerId,
  installmentNum,
  amountPaidNow,
  mop
) => {
  const customer = window.allCustomers.active.find((c) => c.id === customerId);
  if (!customer) throw new Error("Customer not found");
  if (amountPaidNow <= 0) throw new Error("Payment amount must be positive.");

  const updatedSchedule = JSON.parse(JSON.stringify(customer.paymentSchedule));
  const instIndex = updatedSchedule.findIndex(
    (p) => p.installment === installmentNum
  );
  if (instIndex === -1) throw new Error("Installment not found");

  const installment = updatedSchedule[instIndex];
  const prevPaid = Number(installment.amountPaid || 0);
  const newTotalPaid = Math.max(0, prevPaid + amountPaidNow);
  const pending = Math.max(0, Number(installment.amountDue) - newTotalPaid);

  installment.amountPaid = newTotalPaid;
  installment.pendingAmount = pending;
  installment.paidDate = new Date().toISOString();
  installment.modeOfPayment = mop || installment.modeOfPayment || "Cash"; // Use new MOP or fallback

  if (pending <= 0.001) {
    installment.status = "Paid";
    installment.pendingAmount = 0;
  } else if (newTotalPaid > 0) {
    installment.status = "Pending";
  } else {
    installment.status = "Due";
    installment.paidDate = null;
    installment.modeOfPayment = null;
  }

  await db
    .collection("customers")
    .doc(customerId)
    .update({ paymentSchedule: updatedSchedule });
  await logActivity("PAYMENT_RECEIVED", {
    customerName: customer.name,
    amount: amountPaidNow,
  });

  return { customer, installment };
};