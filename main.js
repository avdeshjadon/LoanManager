// --- main.js ---
// Main entry point, authentication, and event listeners

// Authentication state listener
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

// Main function to initialize all event listeners
const initializeEventListeners = () => {
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

  // --- Datepicker Listeners (REMOVED Lines 28-97) ---

  // Main body click listener (delegated)
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
      getEl("reset-password-modal").classList.add("show");
      setTimeout(() => {
        getEl("reset-email")?.focus();
      }, 0);
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

    // Loan Type Selection
    if (button && button.id === "select-loan-type-normal") {
      getEl("loan-type-selection").classList.add("hidden");
      getEl("customer-form").classList.remove("hidden");
      getEl("loan-term-type").value = "normal";

      // Show all normal fields
      getEl("collection-frequency-group").style.display = "block";
      getEl("first-collection-date-group").style.display = "block";
      getEl("loan-end-date-group").style.display = "block";
      getEl("installment-preview").classList.remove("hidden");

      setLoanDetailFieldsRequired(true); // Make them required
    }
    if (button && button.id === "select-loan-type-monthly") {
      getEl("loan-type-selection").classList.add("hidden");
      getEl("customer-form").classList.remove("hidden");
      getEl("loan-term-type").value = "monthly";

      // Hide fields not needed for monthly loan
      getEl("collection-frequency-group").style.display = "none";
      getEl("first-collection-date-group").style.display = "none";
      getEl("loan-end-date-group").style.display = "none";
      getEl("installment-preview").classList.add("hidden");

      // Make them not required
      getEl("collection-frequency").removeAttribute("required");
      getEl("first-collection-date").removeAttribute("required");
      getEl("loan-end-date").removeAttribute("required");
    }

    // Monthly Payment Modal Buttons
    if (button && button.id === "pay-interest-only-btn") {
      const mopSelect = document.getElementById("monthly-payment-mop");
      const selectedMop = mopSelect
        ? mopSelect.value
        : getEl("payment-modal")?.dataset.mop || "Cash";
      if (!selectedMop) {
        showToast("error", "Missing MOP", "Select Mode of Payment.");
        return;
      }
      const customerId = getEl("payment-customer-id").value;
      const installmentNum = parseInt(
        getEl("payment-installment-number").value,
        10
      );
      const mop = getEl("payment-modal")?.dataset.mop || "Cash";
      const customer = window.allCustomers.active.find(
        (c) => c.id === customerId
      );
      if (!customer) return;

      // Interest-only for monthly: pay current month's interest only
      const totalInterest = calculateTotalInterest(
        customer.loanDetails.principal,
        customer.loanDetails.interestRate,
        customer.loanDetails.loanGivenDate,
        customer.loanDetails.loanEndDate
      );

      toggleButtonLoading(button, true, "Paying...");
      try {
        // Record interest-only payment against current installment
        await recordPayment(
          customerId,
          installmentNum,
          totalInterest,
          selectedMop
        );

        // Fetch fresh copy to append next month's EMI and extend end date
        const doc = await db.collection("customers").doc(customerId).get();
        if (doc.exists) {
          const fresh = doc.data();
          const freshSchedule = [...fresh.paymentSchedule];
          const inst = freshSchedule.find(
            (p) => p.installment === installmentNum
          );

          if (
            fresh.loanTermType === "monthly" &&
            inst &&
            inst.status === "Pending" &&
            inst.pendingAmount > 0
          ) {
            // Current installment now has only principal pending.
            // Generate next month due date based on current installment due date.
            const prevDueDate = new Date(inst.dueDate);
            const nextDueDate = addMonthsPreserveAnchor(prevDueDate, 1);

            // Compute ONE MONTH interest using the existing function over [prevDueDate -> nextDueDate]
            const oneMonthInterest = calculateTotalInterest(
              fresh.loanDetails.principal,
              fresh.loanDetails.interestRate,
              prevDueDate,
              nextDueDate
            );

            // Next month's EMI = principal (still pending) + next month's interest
            const nextAmountDue = +(
              inst.pendingAmount + oneMonthInterest
            ).toFixed(2);

            const yyyy = nextDueDate.getFullYear();
            const mm = String(nextDueDate.getMonth() + 1).padStart(2, "0");
            const dd = String(nextDueDate.getDate()).padStart(2, "0");

            freshSchedule.push({
              installment: freshSchedule.length + 1,
              amountDue: nextAmountDue,
              amountPaid: 0,
              pendingAmount: nextAmountDue,
              status: "Due",
              paidDate: null,
              modeOfPayment: null,
              dueDate: `${yyyy}-${mm}-${dd}`,
            });

            // Also extend loan end date to this new due date so Loan Summary and totals reflect the extra month
            const newLoanEndDate = formatForInput({ id: "any" }, nextDueDate);

            await db.collection("customers").doc(customerId).update({
              paymentSchedule: freshSchedule,
              "loanDetails.installments": freshSchedule.length,
              "loanDetails.loanEndDate": newLoanEndDate,
            });

            showToast(
              "success",
              "Next EMI Added",
              "Next month's EMI generated and loan extended by one month."
            );
          }
        }

        showToast("success", "Interest Paid", "Interest payment recorded.");
        getEl("payment-modal").classList.remove("show");
        await loadAndRenderAll();
        showCustomerDetails(customerId);
      } catch (e) {
        showToast("error", "Payment Failed", e.message);
      } finally {
        toggleButtonLoading(button, false);
      }
    }

    if (button && button.id === "pay-full-monthly-btn") {
      const mopSelect = document.getElementById("monthly-payment-mop");
      const selectedMop = mopSelect
        ? mopSelect.value
        : getEl("payment-modal")?.dataset.mop || "Cash";
      if (!selectedMop) {
        showToast("error", "Missing MOP", "Select Mode of Payment.");
        return;
      }
      const customerId = getEl("payment-customer-id").value;
      const installmentNum = parseInt(
        getEl("payment-installment-number").value,
        10
      );
      const mop = getEl("payment-modal")?.dataset.mop || "Cash";
      const customer = window.allCustomers.active.find(
        (c) => c.id === customerId
      );
      if (!customer) return;
      const installment = customer.paymentSchedule.find(
        (p) => p.installment === installmentNum
      );
      const outstandingAmount = installment?.pendingAmount || 0;
      console.log("[FULL MONTHLY PAY]", {
        customerId,
        installmentNum,
        outstandingAmount,
        mop,
      });

      toggleButtonLoading(button, true, "Paying...");
      try {
        await recordPayment(
          customerId,
          installmentNum,
          outstandingAmount,
          selectedMop
        );
        // ...existing code...
      } catch (e) {
        // ...existing code...
      } finally {
        // ...existing code...
      }
    }

    // Undo button handler (normal + monthly)
    if (button && button.classList.contains("undo-payment-btn")) {
      const cid = button.dataset.id;
      const instNum = parseInt(button.dataset.installment, 10);
      toggleButtonLoading(button, true, "Undo...");
      try {
        await undoInstallmentPayment(cid, instNum);
        await loadAndRenderAll();
        showCustomerDetails(cid);
      } catch (err) {
        showToast("error", "Undo Failed", err.message);
      } finally {
        toggleButtonLoading(button, false);
      }
    }

    // Open Pay/Edit modal (normal + monthly)
    if (button && button.classList.contains("record-payment-btn")) {
      const customerId = button.dataset.id;
      const installmentNum = parseInt(button.dataset.installment, 10);
      const customer = window.allCustomers.active.find(
        (c) => c.id === customerId
      );
      if (!customer) return;
      const installment = customer.paymentSchedule.find(
        (p) => p.installment === installmentNum
      );
      if (!installment) return;
      safeSetValue("payment-customer-id", customerId);
      safeSetValue("payment-installment-number", installmentNum);
      const modal = getEl("payment-modal");
      if (!modal) return;
      safeSetText(modal.querySelector(".payment-customer-name"), customer.name);
      safeSetText(
        modal.querySelector(".payment-customer-avatar"),
        customer.name.charAt(0).toUpperCase()
      );
      safeSetText("payment-installment-display", String(installmentNum));
      const defaultMop =
        installment.modeOfPayment ||
        customer.loanDetails?.modeOfPayment ||
        "Cash";

      if (customer.loanTermType === "monthly") {
        getEl("normal-payment-body").style.display = "none";
        getEl("normal-payment-footer").classList.add("hidden");
        getEl("monthly-payment-body").style.display = "block";
        getEl("monthly-payment-footer").classList.remove("hidden");
        const totalInterest = calculateTotalInterest(
          customer.loanDetails.principal,
          customer.loanDetails.interestRate,
          customer.loanDetails.loanGivenDate,
          customer.loanDetails.loanEndDate
        );
        const principal = customer.loanDetails.principal;
        const outstanding = installment.pendingAmount;
        safeSetText("monthly-payment-principal", formatCurrency(principal));
        safeSetText("monthly-payment-interest", formatCurrency(totalInterest));
        safeSetText("payment-due-display", formatCurrency(outstanding));
        modal.dataset.mop = defaultMop;
        if (installment.status === "Pending") {
          const interestBtn = getEl("pay-interest-only-btn");
          if (interestBtn) {
            interestBtn.disabled = true;
            interestBtn.textContent = "Interest Already Paid";
          }
          safeSetText("pay-interest-only-btn-subtext", "");
          safeSetText(
            "pay-full-monthly-btn-subtext",
            formatCurrency(outstanding)
          );
        } else {
          safeSetText(
            "pay-interest-only-btn-subtext",
            formatCurrency(totalInterest)
          );
          safeSetText(
            "pay-full-monthly-btn-subtext",
            formatCurrency(outstanding)
          );
        }
        // Ensure monthly MOP select reflects default
        const monthlyMopSel = document.getElementById("monthly-payment-mop");
        if (monthlyMopSel) monthlyMopSel.value = defaultMop;
      } else {
        getEl("normal-payment-body").style.display = "block";
        getEl("normal-payment-footer").classList.remove("hidden");
        getEl("monthly-payment-body").style.display = "none";
        getEl("monthly-payment-footer").classList.add("hidden");
        const existingPaid = Number(installment.amountPaid || 0);
        const computedPending =
          installment.pendingAmount != null
            ? Number(installment.pendingAmount)
            : Math.max(0, Number(installment.amountDue) - existingPaid);
        safeSetText("payment-due-display", formatCurrency(computedPending));
        const paymentAmountInput = getEl("payment-amount");
        if (paymentAmountInput) {
          paymentAmountInput.value = "";
          paymentAmountInput.setAttribute("max", computedPending);
        }
        const paymentMopInput = getEl("payment-mop");
        if (paymentMopInput) paymentMopInput.value = defaultMop;
        const updatePendingDisplay = () => {
          const payNow = parseFloat(paymentAmountInput?.value || "0") || 0;
          const pendingAmount = Math.max(0, computedPending - payNow);
          const pendingContainer = modal.querySelector(
            ".payment-pending-display"
          );
          if (!pendingContainer) return;
          if (pendingAmount > 0.001) {
            safeSetText(
              pendingContainer.querySelector(".payment-pending-value"),
              formatCurrency(pendingAmount)
            );
            pendingContainer.style.display = "block";
          } else {
            pendingContainer.style.display = "none";
          }
        };
        if (paymentAmountInput)
          paymentAmountInput.oninput = updatePendingDisplay;
        const payFullBtn = getEl("pay-full-btn");
        if (payFullBtn && paymentAmountInput) {
          payFullBtn.onclick = () => {
            paymentAmountInput.value = computedPending;
            updatePendingDisplay();
            paymentAmountInput.focus();
          };
        }
        updatePendingDisplay();
      }
      modal.classList.add("show");
    }

    if (button && button.id === "delete-activity-btn") {
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
    } else if (button.classList.contains("restore-customer-btn")) {
      const customerId = button.dataset.id;
      const customer = window.allCustomers.settled.find(
        (c) => c.id === customerId
      );
      const customerName = customer ? customer.name : "this loan";
      const financeCount = customer ? customer.financeCount || 1 : "";

      showConfirmation(
        "Restore Loan?",
        `This will move Finance ${financeCount} for ${customerName} from 'Settled' back to 'Active'. Are you sure?`,
        async () => {
          await restoreLoanById(customerId);
        }
      );
    } else if (button.classList.contains("delete-customer-btn")) {
      const customerId = button.dataset.id;
      showConfirmation(
        "Delete Customer?",
        "This will permanently delete this customer, their KYC files, and all related activity logs. This action cannot be undone.",
        async () => {
          try {
            await deleteSingleCustomerCascade(customerId);
            const res = await deleteSingleCustomerCascade(customerId);
            if (res && res.name) {
              try {
                await renumberActiveLoansByCustomerName(res.name);
              } catch (e) {
                console.warn("Renumbering after delete failed:", e);
              }
            }
            showToast(
              "success",
              "Customer Deleted",
              "All customer data and files have been removed."
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
        "This will permanently delete ALL settled and refinanced accounts, their KYC files, and related activity logs. This action cannot be undone. Are you sure?",
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
            for (const doc of querySnapshot.docs) {
              const data = doc.data();
              try {
                await deleteSingleCustomerCascade(doc.id, data);
              } catch (e) {
                console.warn(
                  "Failed to delete settled record:",
                  doc.id,
                  e.message || e
                );
              }
            }
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
      getEl("customer-form").classList.add("hidden");
      getEl("loan-type-selection").classList.remove("hidden"); // Show loan type selection
      getEl("customer-id").value = "";
      getEl("customer-form-modal-title").textContent = "Add New Customer";

      // MODIFIED: Use new date function
      const lgd = getEl("loan-given-date");
      if (lgd) lgd.value = formatForDateInput(new Date());

      setAutomaticFirstDate();

      getEl("loan-mop").value = "Cash";

      getEl("personal-info-fields").style.display = "block";
      getEl("kyc-info-fields").style.display = "block";
      getEl("loan-details-fields").style.display = "block";
      if (typeof setLoanDetailFieldsRequired === "function")
        setLoanDetailFieldsRequired(true);
      getEl("installment-preview").classList.add("hidden");

      document.querySelectorAll(".file-input-label span").forEach((span) => {
        span.textContent = "Choose a file...";
        const lbl = span.closest(".file-input-label");
        if (lbl) lbl.title = "";
      });
      getEl("customer-form-modal").classList.add("show");

      // ==========================================
      // == NEW CODE BLOCK FOR PDF/WHATSAPP STARTS ==
      // ==========================================
    } else if (button.id === "generate-pdf-btn") {
      const customerId = button.dataset.id;
      if (!customerId) return;

      toggleButtonLoading(button, true, "Generating...");
      try {
        // This function is expected to be in your 'pdf-generator.js' file
        if (typeof generateAndDownloadPDF === "function") {
          await generateAndDownloadPDF(customerId);
        } else {
          // This error will show if pdf-generator.js is missing or failed to load
          throw new Error(
            "PDF generator function (generateAndDownloadPDF) is not loaded."
          );
        }
      } catch (err) {
        console.error("PDF Generation Error:", err);
        showToast(
          "error",
          "PDF Failed",
          err.message || "Could not generate PDF."
        );
      } finally {
        toggleButtonLoading(button, false);
      }
    } else if (button.id === "send-whatsapp-btn") {
      const customerId = button.dataset.id;
      if (!customerId) return;

      const customer = [
        ...window.allCustomers.active,
        ...window.allCustomers.settled,
      ].find((c) => c.id === customerId);
      if (!customer) {
        showToast("error", "Error", "Customer not found.");
        return;
      }

      // This function is in your 'utils.js' file
      if (typeof openWhatsApp === "function") {
        await openWhatsApp(customer);
      } else {
        showToast("error", "Error", "WhatsApp function not found.");
      }

      // ==========================================
      // == NEW CODE BLOCK FOR PDF/WHATSAPP ENDS ==
      // ==========================================
    } else if (button.id === "edit-customer-info-btn") {
      const customer = [
        ...window.allCustomers.active,
        ...window.allCustomers.settled,
      ].find((c) => c.id === button.dataset.id);
      if (!customer) return;

      // HIDE loan type selection
      getEl("loan-type-selection").classList.add("hidden");
      getEl("customer-form").classList.remove("hidden");

      getEl("customer-form").reset();
      getEl("customer-id").value = customer.id;
      getEl("loan-term-type").value = customer.loanTermType || "normal";

      // Personal fields
      getEl("customer-name").value = customer.name;
      getEl("customer-phone").value = customer.phone || "";

      // MODIFIED: Use new date function
      getEl("customer-dob").value = formatForDateInput(customer.dob || "");

      getEl("customer-father-name").value = customer.fatherName || "";
      getEl("customer-whatsapp").value = customer.whatsapp || "";
      getEl("customer-address").value = customer.address || "";

      getEl("customer-aadhar-number").value = customer.aadharNumber || "";
      getEl("customer-pan-number").value = (
        customer.panNumber || ""
      ).toUpperCase();
      getEl("customer-bank-name").value = customer.bankName || "";
      getEl("customer-account-number").value = customer.accountNumber || "";
      getEl("customer-ifsc").value = customer.ifsc || "";

      // Mode of payment default
      getEl("loan-mop").value = customer.loanDetails?.modeOfPayment || "Cash";

      // Check if loan details are editable (no installment paid)
      const loanEditable = canEditLoanDetailsForCustomer(customer);

      // NEW: Check loanTermType
      if (customer.loanTermType === "monthly") {
        // Hide fields not needed for monthly loan
        getEl("collection-frequency-group").style.display = "none";
        getEl("first-collection-date-group").style.display = "none";
        getEl("loan-end-date-group").style.display = "none";
        getEl("installment-preview").classList.add("hidden");
        // Make them not required
        getEl("collection-frequency").removeAttribute("required");
        getEl("first-collection-date").removeAttribute("required");
        getEl("loan-end-date").removeAttribute("required");
      } else {
        // Show all normal fields
        getEl("collection-frequency-group").style.display = "block";
        getEl("first-collection-date-group").style.display = "block";
        getEl("loan-end-date-group").style.display = "block";
        getEl("installment-preview").classList.remove("hidden");
        setLoanDetailFieldsRequired(true);
      }

      // Show or hide loan detail fields based on editability
      if (loanEditable && customer.loanDetails) {
        // Show and populate loan detail fields
        getEl("loan-details-fields").style.display = "block";

        // Populate loan details fields
        getEl("principal-amount").value = customer.loanDetails.principal || "";
        getEl("interest-rate-modal").value =
          customer.loanDetails.interestRate || "";
        getEl("collection-frequency").value =
          customer.loanDetails.frequency || "monthly";

        // MODIFIED: Use new date function for all date fields
        getEl("first-collection-date").value = formatForDateInput(
          customer.loanDetails.firstCollectionDate || ""
        );
        getEl("loan-end-date").value = formatForDateInput(
          customer.loanDetails.loanEndDate || ""
        );
        getEl("loan-given-date").value = formatForDateInput(
          customer.loanDetails.loanGivenDate || ""
        );
      } else {
        // Hide loan detail fields if not editable
        getEl("loan-details-fields").style.display = "none";
        if (typeof setLoanDetailFieldsRequired === "function")
          setLoanDetailFieldsRequired(false);

        // Notify user why loan details can't be edited
        if (!loanEditable) {
          showToast(
            "error",
            "Cannot Edit Loan Details",
            "Loan details cannot be edited because at least one installment has already been paid."
          );
        }
      }

      getEl("personal-info-fields").style.display = "block";
      getEl("kyc-info-fields").style.display = "block";
      getEl("customer-form-modal-title").textContent = "Edit Customer Info";
      getEl("customer-details-modal").classList.remove("show");
      getEl("customer-form-modal").classList.add("show");
    } else if (button.id === "settle-loan-btn") {
      const currentLoanId = button.dataset.id;
      const currentCustomer = window.allCustomers.active.find(
        (c) => c.id === currentLoanId
      );
      if (!currentCustomer) return;

      // NEW: Don't show modal for monthly loan
      if (currentCustomer.loanTermType === "monthly") {
        showConfirmation(
          `Settle Loan?`,
          `This will move this Monthly Loan to the 'Settled' list. Are you sure?`,
          () => {
            settleLoanById(currentLoanId);
          }
        );
        return;
      }

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
        showConfirmation(
          "Settle This Loan?",
          "This will move the selected loan to Settled. Continue?",
          () => settleLoanById(loanIdToSettle)
        );
      } else {
        showToast("error", "No Selection", "Please select a loan to settle.");
      }
    } else if (button.id === "settle-all-btn") {
      const firstInput = document.querySelector('input[name="settle-loan"]');
      if (!firstInput) {
        showToast(
          "error",
          "No Loans",
          "No active loans found for this customer."
        );
        return;
      }
      const inputs = Array.from(
        document.querySelectorAll('input[name="settle-loan"]')
      );
      const loanIds = inputs.map((r) => r.value);
      if (loanIds.length === 0) {
        showToast(
          "error",
          "No Loans",
          "No active loans found for this customer."
        );
        return;
      }
      showConfirmation(
        "Settle ALL Loans?",
        "This will move all selected customer's active loans to Settled. Continue?",
        async () => {
          try {
            for (const id of loanIds) {
              await settleLoanById(id);
            }
            showToast(
              "success",
              "All Settled",
              "All active loans for the customer have been settled."
            );
          } catch (err) {
            showToast("error", "Failed", err.message || String(err));
          }
        }
      );
    } else if (button.id === "add-new-loan-btn") {
      const customerId = button.dataset.id;
      const customer = window.allCustomers.active.find(
        (c) => c.id === customerId
      );
      if (!customer) return;

      getEl("new-loan-customer-id").value = customerId;
      getEl("new-loan-form").reset();

      // MODIFIED: Use new date function
      const nlgd = getEl("new-loan-given-date");
      if (nlgd) nlgd.value = formatForDateInput(new Date());

      const freq = getEl("new-loan-frequency").value;
      let base = new Date();
      if (freq === "daily") base.setDate(base.getDate() + 1);
      else if (freq === "weekly") base.setDate(base.getDate() + 7);
      else if (freq === "monthly") base.setMonth(base.getMonth() + 1);

      // MODIFIED: Use new date function
      getEl("new-loan-start-date").value = formatForDateInput(base);

      getEl("new-loan-mop").value = "Cash";

      getEl("new-loan-installment-preview").classList.add("hidden");
      getEl("new-loan-modal").classList.add("show");
    } else if (button.id === "export-active-btn") {
      exportToExcel(window.allCustomers.active, "Active_Customers_Report.xlsx");
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
        showToast("error", "No File", "Please choose a backup file to import.");
        return;
      }
      showConfirmation(
        "Overwrite All Data?",
        "Importing this file will permanently delete ALL your current customers and replacing them with the data from the backup. This cannot be undone. Are you sure you want to proceed?",
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
              existingCustomers.docs.forEach((doc) => batch.delete(doc.ref));
              [...backup.customers.active, ...backup.customers.settled].forEach(
                (cust) => {
                  const newDocRef = db.collection("customers").doc();
                  const newCustData = {
                    ...cust,
                    owner: currentUser.uid,
                    createdAt: new Date(),
                  };
                  delete newCustData.id;
                  batch.set(newDocRef, newCustData);
                }
              );
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
  });

  // Main body form submission listener (delegated)
  document.body.addEventListener("submit", async (e) => {
    e.preventDefault();
    const form = e.target;
    if (form.id === "reset-password-form") {
      const btn = getEl("reset-password-submit");
      toggleButtonLoading(btn, true, "Sending...");
      try {
        const email = getEl("reset-email").value.trim();
        await auth.sendPasswordResetEmail(email);
        showToast(
          "success",
          "Email Sent",
          "Password reset link sent to your email."
        );
        getEl("reset-password-modal").classList.remove("show");
      } catch (error) {
        showToast("error", "Error", error.message);
      } finally {
        toggleButtonLoading(btn, false);
      }
      return;
    }
    if (form.id === "customer-form") {
      if (!form.checkValidity()) {
        form.reportValidity();
        return;
      }
      let reopenAfterSaveId = null;
      const id = getEl("customer-id").value;
      const saveBtn = getEl("customer-modal-save");
      toggleButtonLoading(saveBtn, true, id ? "Updating..." : "Saving...");
      if (!id) {
        const p = parseFloat(getEl("principal-amount").value);
        const r = parseFloat(getEl("interest-rate-modal").value);

        if (isNaN(p) || isNaN(r) || p <= 0) {
          showToast(
            "error",
            "Invalid Input",
            "Please enter a valid Principal and Interest rate."
          );
          toggleButtonLoading(saveBtn, false); // Button ko re-enable karo
          return; // Function se bahar nikal jao
        }
      }

      try {
        const customerData = {
          name: getEl("customer-name").value,
          phone: getEl("customer-phone").value,
          dob: (() => {
            const raw = getEl("customer-dob").value;
            if (!raw) return "";
            // Native date input gives YYYY-MM-DD, parse it correctly
            const d = parseDateFlexible(raw);
            // Store as DD-MM-YYYY
            return isNaN(+d) ? raw : formatForInput({ id: "any" }, d);
          })(),
          fatherName: getEl("customer-father-name").value,
          address: getEl("customer-address").value,
          whatsapp: getEl("customer-whatsapp").value,
          aadharNumber: getEl("customer-aadhar-number").value.trim(),
          panNumber: getEl("customer-pan-number").value.trim().toUpperCase(),
          bankName: getEl("customer-bank-name")?.value || "",
          accountNumber: getEl("customer-account-number")?.value || "",
          ifsc: getEl("customer-ifsc")?.value || "",
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
        const [aadharUrlFront, aadharUrlBack, panUrl, picUrl, bankDetailsUrl] =
          await Promise.all([
            uploadFile("customer-aadhar-file-front", "aadhar-front"),
            uploadFile("customer-aadhar-file-back", "aadhar-back"),
            uploadFile("customer-pan-file", "pan"),
            uploadFile("customer-pic-file", "picture"),
            uploadFile("customer-bank-file", "bank"),
          ]);

        customerData.kycDocs = {
          ...(aadharUrlFront && { aadharUrlFront }),
          ...(aadharUrlBack && { aadharUrlBack }),
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
            dob: updatePayload.dob,
            fatherName: updatePayload.fatherName,
            address: updatePayload.address,
            whatsapp: updatePayload.whatsapp,
            aadharNumber: getEl("customer-aadhar-number").value.trim(),
            panNumber: getEl("customer-pan-number").value.trim().toUpperCase(),
            bankName: getEl("customer-bank-name")?.value || "",
            accountNumber: getEl("customer-account-number")?.value || "",
            ifsc: getEl("customer-ifsc")?.value || "",
          };

          // File uploads
          if (aadharUrlFront)
            finalUpdate["kycDocs.aadharUrlFront"] = aadharUrlFront;
          if (aadharUrlBack)
            finalUpdate["kycDocs.aadharUrlBack"] = aadharUrlBack;
          if (panUrl) finalUpdate["kycDocs.panUrl"] = panUrl;
          if (picUrl) finalUpdate["kycDocs.picUrl"] = picUrl;
          if (bankDetailsUrl)
            finalUpdate["kycDocs.bankDetailsUrl"] = bankDetailsUrl;

          // Always update mode of payment
          finalUpdate["loanDetails.modeOfPayment"] = getEl("loan-mop").value;

          // Check if loan details are editable (no installment paid)
          const existingDoc = await db.collection("customers").doc(id).get();
          const existingData = existingDoc.exists ? existingDoc.data() : null;
          const canEditLoan = canEditLoanDetailsForCustomer(existingData);

          // Only include loan detail updates if they are editable
          if (
            canEditLoan &&
            existingData &&
            existingData.loanDetails &&
            existingData.paymentSchedule &&
            existingData.loanTermType !== "monthly" // Don't allow edits for monthly loan details
          ) {
            // Get all the updated loan details
            const updatedPrincipal = parseFloat(
              getEl("principal-amount")?.value
            );
            const updatedRate = parseFloat(getEl("interest-rate-modal")?.value);
            const updatedFrequency = getEl("collection-frequency")?.value;
            // Native date input gives YYYY-MM-DD, parse it correctly
            const updatedFirstDate = getEl("first-collection-date")?.value;
            const updatedEndDate = getEl("loan-end-date")?.value;
            const updatedLoanGivenDate = getEl("loan-given-date")?.value;

            // Check if we have all values needed to recalculate schedule
            if (
              !isNaN(updatedPrincipal) &&
              !isNaN(updatedRate) &&
              updatedFrequency &&
              updatedFirstDate &&
              updatedEndDate &&
              updatedLoanGivenDate
            ) {
              // Calculate new installments based on updated date range and frequency
              const newNBase = calculateInstallments(
                updatedFirstDate,
                updatedEndDate,
                updatedFrequency
              );

              // Calculate new total repayable amount
              const newTotalInterest = calculateTotalInterest(
                updatedPrincipal,
                updatedRate,
                updatedLoanGivenDate,
                updatedEndDate
              );
              const newTotalRepayable = updatedPrincipal + newTotalInterest;

              // Get the installment amount (either custom or standard)
              let newSchedule;
              const customAmtInput = getEl("custom-installment-amount");
              let chosenN = newNBase;

              if (customAmtInput && customAmtInput.value) {
                const minInstallment = +(newTotalRepayable / newNBase).toFixed(
                  2
                );
                let chosenAmt = parseFloat(customAmtInput.value);
                if (
                  !chosenAmt ||
                  isNaN(chosenAmt) ||
                  chosenAmt < minInstallment
                ) {
                  chosenAmt = minInstallment;
                }
                chosenN = Math.max(1, Math.ceil(newTotalRepayable / chosenAmt));
                newSchedule = generateScheduleWithInstallmentAmount(
                  +totalRepayable.toFixed(2),
                  +chosenAmt.toFixed(2)
                );
              } else {
                // Use standard equal installments
                newSchedule = generateSimpleInterestSchedule(
                  +totalRepayable.toFixed(2),
                  newNBase
                );
              }

              // Assign due dates to the new schedule
              let currentDate = parseDateFlexible(updatedFirstDate);
              newSchedule.forEach((inst, index) => {
                if (index > 0) {
                  if (updatedFrequency === "daily") {
                    currentDate.setDate(currentDate.getDate() + 1);
                  } else if (updatedFrequency === "weekly") {
                    currentDate.setDate(currentDate.getDate() + 7);
                  } else if (updatedFrequency === "monthly") {
                    const base = parseDateFlexible(updatedFirstDate);
                    currentDate = addMonthsPreserveAnchor(base, index);
                  }
                }
                const yyyy = currentDate.getFullYear();
                const mm = String(currentDate.getMonth() + 1).padStart(2, "0");
                const dd = String(currentDate.getDate()).padStart(2, "0");
                inst.dueDate = `${yyyy}-${mm}-${dd}`;
              });

              // Update all loan details and the payment schedule
              finalUpdate["loanDetails.principal"] = updatedPrincipal;
              finalUpdate["loanDetails.interestRate"] = updatedRate;
              finalUpdate["loanDetails.frequency"] = updatedFrequency;
              finalUpdate["loanDetails.firstCollectionDate"] = formatForInput(
                { id: "any" },
                parseDateFlexible(updatedFirstDate)
              ); // Store as DD-MM-YYYY
              finalUpdate["loanDetails.loanEndDate"] = formatForInput(
                { id: "any" },
                parseDateFlexible(updatedEndDate)
              ); // Store as DD-MM-YYYY
              finalUpdate["loanDetails.loanGivenDate"] = formatForInput(
                { id: "any" },
                parseDateFlexible(updatedLoanGivenDate)
              ); // Store as DD-MM-YYYY
              finalUpdate["loanDetails.installments"] = chosenN;
              finalUpdate["paymentSchedule"] = newSchedule;

              // Update hidden fields if they exist
              const instEl = getEl("chosen-n-installments");
              if (instEl && instEl.value) {
                const instNum = parseInt(instEl.value, 10);
                if (!isNaN(instNum)) {
                  finalUpdate["loanDetails.installments"] = instNum;
                }
              }
            } else {
              // Individual updates if we don't have enough data for full recalculation
              if (!isNaN(updatedPrincipal))
                finalUpdate["loanDetails.principal"] = updatedPrincipal;
              if (!isNaN(updatedRate))
                finalUpdate["loanDetails.interestRate"] = updatedRate;
              if (updatedFrequency)
                finalUpdate["loanDetails.frequency"] = updatedFrequency;
              if (updatedFirstDate)
                finalUpdate["loanDetails.firstCollectionDate"] = formatForInput(
                  { id: "any" },
                  parseDateFlexible(updatedFirstDate)
                );
              if (updatedEndDate)
                finalUpdate["loanDetails.loanEndDate"] = formatForInput(
                  { id: "any" },
                  parseDateFlexible(updatedEndDate)
                );
              if (updatedLoanGivenDate)
                finalUpdate["loanDetails.loanGivenDate"] = formatForInput(
                  { id: "any" },
                  parseDateFlexible(updatedLoanGivenDate)
                );

              const instEl = getEl("chosen-n-installments");
              if (instEl && instEl.value) {
                const instNum = parseInt(instEl.value, 10);
                if (!isNaN(instNum))
                  finalUpdate["loanDetails.installments"] = instNum;
              }
            }
          }

          await db.collection("customers").doc(id).update(finalUpdate);
          reopenAfterSaveId = id;
          showToast(
            "success",
            "Customer Updated",
            "Details saved successfully."
          );
          // Wait for refresh before showing details after update
          getEl("customer-form-modal").classList.remove("show");
          await loadAndRenderAll();
          if (reopenAfterSaveId) {
            showCustomerDetails(reopenAfterSaveId);
          }
        } else {
          // NEW CUSTOMER
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
          const loanTermType = getEl("loan-term-type").value;
          customerData.loanTermType = loanTermType;

          const loanGivenDate = getEl("loan-given-date")?.value || new Date();
          const lgdDate =
            typeof loanGivenDate === "string"
              ? parseDateFlexible(loanGivenDate)
              : loanGivenDate;
          const lgdStr = formatForInput({ id: "any" }, lgdDate); // Store as DD-MM-YYYY

          if (loanTermType === "monthly") {
            // --- MONTHLY LOAN LOGIC ---
            const endDate = addMonthsPreserveAnchor(lgdDate, 1);
            const endDateStr = formatForInput({ id: "any" }, endDate); // Store as DD-MM-YYYY
            const firstDateStr = endDateStr; // Only one payment, at the end

            const totalInterest = calculateTotalInterest(
              p,
              r,
              lgdStr,
              endDateStr
            );
            const totalRepayable = p + totalInterest;

            customerData.loanDetails = {
              principal: p,
              interestRate: r,
              installments: 1,
              frequency: "monthly", // Store this for consistency, though UI is hidden
              loanGivenDate: lgdStr,
              firstCollectionDate: firstDateStr,
              loanEndDate: endDateStr,
              type: "simple_interest", // Keep this type
              modeOfPayment: getEl("loan-mop").value,
            };

            const yyyy = endDate.getFullYear();
            const mm = String(endDate.getMonth() + 1).padStart(2, "0");
            const dd = String(endDate.getDate()).padStart(2, "0");

            customerData.paymentSchedule = [
              {
                installment: 1,
                amountDue: +totalRepayable.toFixed(2),
                amountPaid: 0,
                pendingAmount: +totalRepayable.toFixed(2),
                status: "Due",
                paidDate: null,
                modeOfPayment: null,
                dueDate: `${yyyy}-${mm}-${dd}`,
              },
            ];
          } else {
            // --- NORMAL LOAN LOGIC ---
            const freq = getEl("collection-frequency").value;
            // Native date input gives YYYY-MM-DD
            const firstDate = getEl("first-collection-date").value;
            const endDate = getEl("loan-end-date").value;

            const nBase = calculateInstallments(firstDate, endDate, freq);

            if (isNaN(p) || isNaN(r) || isNaN(nBase) || !firstDate)
              throw new Error("Please fill all loan detail fields correctly.");

            const totalRepayable =
              p + calculateTotalInterest(p, r, lgdStr, endDate);

            const customAmtInput = getEl("custom-installment-amount");
            let chosenN = nBase;
            let chosenEndDate = endDate;
            let schedule;
            if (customAmtInput) {
              const minInstallment = +(totalRepayable / nBase).toFixed(2);
              let chosenAmt = parseFloat(customAmtInput.value);
              if (!chosenAmt || isNaN(chosenAmt) || chosenAmt < minInstallment)
                chosenAmt = minInstallment;
              chosenN = Math.max(1, Math.ceil(totalRepayable / chosenAmt));
              chosenEndDate = computeEndDateFromInstallments(
                firstDate,
                chosenN,
                freq
              );
              schedule = generateScheduleWithInstallmentAmount(
                +totalRepayable.toFixed(2),
                +chosenAmt.toFixed(2)
              );
            }

            customerData.loanDetails = {
              principal: p,
              interestRate: r,
              installments: chosenN,
              frequency: freq,
              loanGivenDate: lgdStr,
              firstCollectionDate: formatForInput(
                { id: "any" },
                parseDateFlexible(firstDate)
              ), // Store as DD-MM-YYYY
              loanEndDate: formatForInput(
                { id: "any" },
                parseDateFlexible(endDate)
              ), // Store as DD-MM-YYYY
              type: "simple_interest",
              modeOfPayment: getEl("loan-mop").value,
            };

            let paymentSchedule = schedule
              ? schedule
              : generateSimpleInterestSchedule(
                  +totalRepayable.toFixed(2),
                  nBase
                );

            let currentDate = parseDateFlexible(firstDate);
            paymentSchedule.forEach((inst, index) => {
              if (index > 0) {
                if (freq === "daily") {
                  currentDate.setDate(currentDate.getDate() + 1);
                } else if (freq === "weekly") {
                  currentDate.setDate(currentDate.getDate() + 7);
                } else if (freq === "monthly") {
                  const base = parseDateFlexible(firstDate);
                  currentDate = addMonthsPreserveAnchor(base, index);
                }
              }
              const yyyy = currentDate.getFullYear();
              const mm = String(currentDate.getMonth() + 1).padStart(2, "0");
              const dd = String(currentDate.getDate()).padStart(2, "0");
              inst.dueDate = `${yyyy}-${mm}-${dd}`;
            });

            customerData.paymentSchedule = paymentSchedule;
          }

          customerData.owner = currentUser.uid;
          customerData.createdAt =
            firebase.firestore.FieldValue.serverTimestamp();
          customerData.status = "active";
          customerData.financeCount = 1;

          toggleButtonLoading(saveBtn, true, "Saving Customer...");
          const docRef = await db.collection("customers").add(customerData);
          reopenAfterSaveId = docRef.id;
          await logActivity("NEW_LOAN", {
            customerName: customerData.name,
            amount: p,
            financeCount: customerData.financeCount,
          });
          showToast("success", "Customer Added", "New loan account created.");

          // ** FIX IS HERE: Wait for data reload BEFORE showing details **
          getEl("new-loan-modal").classList.remove("show");
          await loadAndRenderAll(); // Wait for the refresh to complete
          showCustomerDetails(docRef.id); // Now uses updated data
        }
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
        const amountPaidNow = parseFloat(getEl("payment-amount").value) || 0;
        const mop = getEl("payment-mop").value;

        if (!mop) {
          throw new Error("Please select a mode of payment.");
        }

        const { customer, installment } = await recordPayment(
          customerId,
          installmentNum,
          amountPaidNow,
          mop
        );

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

      const totalInterest = calculateTotalInterestByTerm(p, r, n, freq);
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

        const nBase = calculateInstallments(firstDate, endDate, freq);

        if (isNaN(p) || isNaN(r) || isNaN(nBase) || !firstDate || !endDate)
          throw new Error("Please fill all new loan fields correctly.");

        const loanGivenDate =
          getEl("new-loan-given-date")?.value ||
          getEl("loan-given-date")?.value ||
          new Date();
        const lgdDate =
          typeof loanGivenDate === "string"
            ? parseDateFlexible(loanGivenDate)
            : loanGivenDate;
        const lgdStr = formatForInput({ id: "any" }, lgdDate); // Store as DD-MM-YYYY

        const totalRepayable =
          p + calculateTotalInterest(p, r, lgdStr, endDate);

        const customAmtInput = getEl("new-loan-custom-installment-amount");
        let chosenN = nBase;
        let chosenEndDate = endDate;
        let schedule;
        if (customAmtInput) {
          const minInstallment = +(totalRepayable / nBase).toFixed(2);
          let chosenAmt = parseFloat(customAmtInput.value);
          if (!chosenAmt || isNaN(chosenAmt) || chosenAmt < minInstallment)
            chosenAmt = minInstallment;
          chosenN = Math.max(1, Math.ceil(totalRepayable / chosenAmt));
          chosenEndDate = computeEndDateFromInstallments(
            firstDate,
            chosenN,
            freq
          );
          schedule = generateScheduleWithInstallmentAmount(
            +totalRepayable.toFixed(2),
            +chosenAmt.toFixed(2)
          );
        }

        const newLoanData = {
          name: baseCustomer.name,
          phone: baseCustomer.phone,
          fatherName: baseCustomer.fatherName,
          address: baseCustomer.address,
          whatsapp: baseCustomer.whatsapp,
          kycDocs: baseCustomer.kycDocs || {},
          bankName: baseCustomer.bankName || "",
          accountNumber: baseCustomer.accountNumber || "",
          ifsc: baseCustomer.ifsc || "",
          owner: currentUser.uid,
          createdAt: firebase.firestore.FieldValue.serverTimestamp(),
          status: "active",
          financeCount: maxFinanceCount + 1,
          loanTermType: "normal", // New loans for existing customers are "normal"
          loanDetails: {
            principal: p,
            interestRate: r,
            installments: chosenN,
            frequency: freq,
            loanGivenDate: formatForInput({ id: "any" }, lgdDate),
            firstCollectionDate: formatForInput(
              { id: "any" },
              parseDateFlexible(firstDate)
            ), // Store as DD-MM-YYYY
            loanEndDate: formatForInput(
              { id: "any" },
              parseDateFlexible(endDate)
            ), // Store as DD-MM-YYYY
            type: "simple_interest",
            modeOfPayment: getEl("new-loan-mop").value,
          },
        };

        const chosenHiddenN = getEl("new-loan-chosen-n-installments");
        const effectiveN =
          chosenHiddenN && chosenHiddenN.value
            ? parseInt(chosenHiddenN.value, 10)
            : nBase;
        let paymentSchedule = schedule
          ? schedule
          : generateSimpleInterestSchedule(
              +totalRepayable.toFixed(2),
              effectiveN
            );

        const chosenHiddenEnd = getEl("new-loan-chosen-end-date");
        if (chosenHiddenEnd && chosenHiddenEnd.value) {
          newLoanData.loanDetails.loanEndDate = formatForInput(
            { id: "any" },
            parseDateFlexible(chosenHiddenEnd.value)
          ); // Store as DD-MM-YYYY
        }

        let currentDate = parseDateFlexible(firstDate);
        paymentSchedule.forEach((inst, index) => {
          if (index > 0) {
            if (freq === "daily") {
              currentDate.setDate(currentDate.getDate() + 1);
            } else if (freq === "weekly") {
              currentDate.setDate(currentDate.getDate() + 7);
            } else if (freq === "monthly") {
              const base = parseDateFlexible(firstDate);
              currentDate = addMonthsPreserveAnchor(base, index);
            }
          }
          const yyyy = currentDate.getFullYear();
          const mm = String(currentDate.getMonth() + 1).padStart(2, "0");
          const dd = String(currentDate.getDate()).padStart(2, "0");
          inst.dueDate = `${yyyy}-${mm}-${dd}`;
        });

        newLoanData.paymentSchedule = paymentSchedule;

        const docRef = await db.collection("customers").add(newLoanData);
        await logActivity("NEW_LOAN", {
          customerName: newLoanData.name,
          amount: p,
          financeCount: newLoanData.financeCount,
        });
        showToast(
          "success",
          "New Loan Added",
          "New loan account created successfully."
        );

        getEl("new-loan-modal").classList.remove("show");
        await loadAndRenderAll();
        showCustomerDetails(docRef.id);
      } catch (e) {
        showToast("error", "Save Failed", e.message);
      } finally {
        toggleButtonLoading(saveBtn, false);
      }
    } else if (form.id === "change-password-form") {
      const btn = getEl("change-password-btn");
      const currentPassword = getEl("current-password").value;
      const newPassword = getEl("new-password").value;
      const confirmPassword = getEl("confirm-password").value;

      if (newPassword !== confirmPassword) {
        showToast("error", "Error", "New passwords do not match.");
        return;
      }
      if (newPassword.length < 6) {
        showToast(
          "error",
          "Error",
          "Password must be at least 6 characters long."
        );
        return;
      }

      toggleButtonLoading(btn, true, "Updating...");
      try {
        const user = auth.currentUser;
        const credential = firebase.auth.EmailAuthProvider.credential(
          user.email,
          currentPassword
        );
        await user.reauthenticateWithCredential(credential);
        await user.updatePassword(newPassword);
        showToast(
          "success",
          "Password Updated",
          "Your password has been changed successfully."
        );
        form.reset();
      } catch (error) {
        showToast("error", "Update Failed", error.message);
      } finally {
        toggleButtonLoading(btn, false);
      }
    }
  });

  // Main body change listener (delegated)
  document.body.addEventListener("change", (e) => {
    const el = e.target;
    if (el.classList.contains("file-input")) {
      const label = el.nextElementSibling;
      const span = label.querySelector("span");
      if (el.files && el.files.length > 0) {
        const fileName = el.files[0].name;
        const truncated = truncateMiddle(fileName, 25);
        span.textContent = truncated;
        label.title = fileName;

        // File size check
        const fileSize = el.files[0].size; // in bytes
        const maxSize = 1 * 1024 * 1024; // 1 MB
        if (fileSize > maxSize) {
          showSizeAlert();
          el.value = ""; // Clear the file input
          span.textContent = "Choose a file...";
          label.title = "";
        }
      } else {
        span.textContent = "Choose a file...";
        label.title = "";
      }
    }

    // Installment calculation triggers
    if (
      [
        "principal-amount",
        "interest-rate-modal",
        "collection-frequency",
        "first-collection-date",
        "loan-end-date",
        "loan-given-date",
      ].includes(el.id)
    ) {
      if (getEl("loan-term-type").value === "normal") {
        updateInstallmentPreview();
      }
    }

    if (el.id === "collection-frequency") {
      setAutomaticFirstDate();
    }

    // New loan calculation triggers
    if (
      [
        "new-loan-principal",
        "new-loan-interest-rate",
        "new-loan-frequency",
        "new-loan-start-date",
        "new-loan-end-date",
        "new-loan-given-date",
      ].includes(el.id)
    ) {
      updateNewLoanInstallmentPreview();
    }

    if (el.id === "new-loan-frequency") {
      setAutomaticNewLoanFirstDate();
    }

    // Import file display
    if (el.id === "import-backup-input") {
      const display = getEl("file-name-display");
      if (el.files && el.files.length > 0) {
        display.textContent = el.files[0].name;
      } else {
        display.textContent = "No file chosen";
      }
    }

    // Dark mode toggle
    if (el.id === "dark-mode-toggle") {
      if (window.toggleDarkMode) window.toggleDarkMode();
    }

    // Customer sort
    if (el.id === "customer-sort-select") {
      activeSortKey = el.value;
      renderActiveCustomerList(window.allCustomers.active, activeSortKey);
    }
  });

  // Main body input listener (delegated)
  document.body.addEventListener("input", (e) => {
    const el = e.target;
    if (el.id === "search-customers") {
      const searchTerm = el.value.toLowerCase().trim();
      const filtered = window.allCustomers.active.filter((c) =>
        c.name.toLowerCase().includes(searchTerm)
      );
      renderActiveCustomerList(filtered, activeSortKey);
    }
  });
};

// Start the application after the DOM is loaded
document.addEventListener("DOMContentLoaded", () => {
  initializeEventListeners();
});
