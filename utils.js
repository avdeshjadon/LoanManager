// --- utils.js ---
// Helper functions, calculations, aur utilities

const parseDateFlexible = (input) => {
  if (input instanceof Date) return input;
  if (typeof input === "number") return new Date(input);
  if (typeof input !== "string") {
    const d = new Date(input);
    return d;
  }
  const s = input.trim();
  let m = s.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const d = new Date(s);
  return d;
};

const addMonthsPreserveAnchor = (date, months) => {
  const y = date.getFullYear();
  const m = date.getMonth();
  const d = date.getDate();
  const target = new Date(y, m + months, 1);
  const lastDay = new Date(
    target.getFullYear(),
    target.getMonth() + 1,
    0
  ).getDate();
  const day = Math.min(d, lastDay);
  return new Date(target.getFullYear(), target.getMonth(), day);
};

const countAnchorMonths = (start, end) => {
  const s = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const e = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  if (e <= s) return 0;
  let months = 1;
  while (e > addMonthsPreserveAnchor(s, months)) {
    months += 1;
  }
  return months;
};

const calculateTotalInterest = (
  principal,
  monthlyRate,
  loanGivenDate,
  loanEndDate
) => {
  if (!principal || !monthlyRate || !loanGivenDate || !loanEndDate) return 0;
  const monthlyRateDecimal = monthlyRate / 100;
  const start = parseDateFlexible(loanGivenDate);
  const end = parseDateFlexible(loanEndDate);
  if (!(start instanceof Date) || isNaN(+start)) return 0;
  if (!(end instanceof Date) || isNaN(+end)) return 0;
  if (end <= start) return 0;

  const monthsBlocks = countAnchorMonths(start, end);
  return principal * monthlyRateDecimal * monthsBlocks;
};

const calculateTotalInterestByTerm = (
  principal,
  monthlyRate,
  numberOfInstallments,
  frequency
) => {
  if (
    !principal ||
    !monthlyRate ||
    !numberOfInstallments ||
    numberOfInstallments <= 0
  )
    return 0;
  const start = new Date();
  const s = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  let end = new Date(s);
  switch (frequency) {
    case "monthly": {
      end = addMonthsPreserveAnchor(s, numberOfInstallments);
      break;
    }
    case "weekly": {
      end.setDate(end.getDate() + numberOfInstallments * 7);
      break;
    }
    case "daily": {
      end.setDate(end.getDate() + numberOfInstallments);
      break;
    }
    default:
      return 0;
  }
  return calculateTotalInterest(principal, monthlyRate, s, end);
};

const formatCurrency = (amount) => {
  const value = Number(amount || 0);
  return `₹${value.toLocaleString("en-IN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}`;
};

// NEW function for native date inputs (YYYY-MM-DD)
const formatForDateInput = (date) => {
  if (!date) return "";
  // parseDateFlexible ka istemal karo taaki yeh DD-MM-YYYY string ko bhi handle kar sake
  const d = parseDateFlexible(date);
  if (isNaN(+d)) return "";
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
};

function openWhatsApp(customer) {
  if (!customer || (!customer.whatsapp && !customer.phone)) {
    alert("Customer's WhatsApp number is not available.");
    return;
  }

  // format phone
  let phone = (customer.whatsapp || customer.phone || "").replace(/\D/g, "");
  if (phone.length === 10) phone = "91" + phone;

  // safe helpers / fallbacks
  const safeNumber = (v) => {
    const n = Number(v || 0);
    return isNaN(n) ? 0 : n;
  };
  const formatCurrencySafe = (amt) => {
    try {
      if (typeof formatCurrency === "function") return formatCurrency(amt);
      const v = Number(amt || 0);
      return `₹${v.toLocaleString("en-IN", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      })}`;
    } catch {
      return `₹${(Number(amt) || 0).toFixed(2)}`;
    }
  };

  const loanDetails = customer.loanDetails || {};
  let totalInterest = 0;
  if (typeof calculateTotalInterest === "function") {
    try {
      totalInterest = calculateTotalInterest(
        safeNumber(loanDetails.principal),
        safeNumber(loanDetails.interestRate),
        loanDetails.loanGivenDate,
        loanDetails.loanEndDate
      );
    } catch {
      totalInterest = 0;
    }
  }

  const totalRepayable =
    safeNumber(loanDetails.principal) + safeNumber(totalInterest);

  let message = `Dear Partner,\n\n`;
  message += `We are pleased to provide a summary of your outstanding balance with Global Finance Consultant as outlined below:\n\n`;
  message += `1. Total Principal Amount:   ${formatCurrencySafe(
    loanDetails.principal
  )}\n`;
  message += `2. Total Interest Amount:     ${formatCurrencySafe(totalInterest)}\n`;
  message += `3. Total Payable Amount:      ${formatCurrencySafe(totalRepayable)}\n\n`;
  message += `We kindly request that all payments be made in accordance with the agreed repayment schedule to avoid any late fees.\n\n`;
  message += `*A late fee of ₹50 per day will apply if payment is not made more than three days after the due date until the payment is received.*\n\n`;
  message += `Best regards,\n`;
  message += `Global Finance Consultant`;

  const whatsappUrl = `https://wa.me/${phone}?text=${encodeURIComponent(
    message
  )}`;
  window.open(whatsappUrl, "_blank");
}

window.openWhatsApp = openWhatsApp;

window.processProfitData = (allCusts) => {
  const data = [];
  const today = new Date();
  today.setHours(23, 59, 59, 999);
  [...allCusts.active, ...allCusts.settled].forEach((customer) => {
    if (customer.loanDetails && customer.paymentSchedule) {
      const totalInterest = calculateTotalInterest(
        customer.loanDetails.principal,
        customer.loanDetails.interestRate,
        customer.loanDetails.loanGivenDate,
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

const pad2 = (n) => String(n).padStart(2, "0");

const formatForInput = (inputEl, date) => {
  const id = inputEl.id || "";
  return `${pad2(date.getDate())}-${pad2(
    date.getMonth() + 1
  )}-${date.getFullYear()}`;
};

// parseFromInput function (lines 320-338) REMOVED

const truncateMiddle = (name, maxChars) => {
  if (!name || typeof name !== "string") return "";
  if (!maxChars || name.length <= maxChars) return name;
  const dot = name.lastIndexOf(".");
  const ext = dot > 0 ? name.slice(dot) : "";
  const base = dot > 0 ? name.slice(0, dot) : name;
  const room = Math.max(3, maxChars - ext.length);
  if (base.length + ext.length <= maxChars) return name;
  const keep = Math.max(2, Math.floor((room - 3) / 2));
  return `${base.slice(0, keep)}...${base.slice(-keep)}${ext}`;
};

const calculateInstallments = (startDateStr, endDateStr, frequency) => {
  if (!startDateStr || !endDateStr) return 0;
  const parseAny = (s) => {
    const dm = s.match(/^(\d{2})-(\d{2})-(\d{4})$/);
    if (dm) return new Date(Number(dm[3]), Number(dm[2]) - 1, Number(dm[1]));
    const ym = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (ym) return new Date(Number(ym[1]), Number(ym[2]) - 1, Number(ym[3]));
    const d = new Date(s);
    return isNaN(+d) ? new Date() : d;
  };
  const startDate = parseAny(startDateStr);
  const endDate = parseAny(endDateStr);

  if (endDate < startDate) {
    throw new Error("End date must be on or after the start date.");
  }

  if (frequency === "monthly") {
    let months;
    months = (endDate.getFullYear() - startDate.getFullYear()) * 12;
    months -= startDate.getMonth();
    months += endDate.getMonth();

    if (endDate.getDate() < startDate.getDate()) {
      months--;
    }
    return months + 1;
  }

  const diffTime = Math.abs(endDate - startDate);
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;

  if (frequency === "daily") {
    return diffDays;
  }
  if (frequency === "weekly") {
    return Math.ceil(diffDays / 7);
  }

  throw new Error("Invalid frequency selected.");
};

const computeEndDateFromInstallments = (startDateStr, n, frequency) => {
  if (!startDateStr || !n || n <= 0) return startDateStr;
  const parseAny = (s) => {
    const dm = s.match(/^(\d{2})-(\d{2})-(\d{4})$/);
    if (dm) return new Date(Number(dm[3]), Number(dm[2]) - 1, Number(dm[1]));
    const ym = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (ym) return new Date(Number(ym[1]), Number(ym[2]) - 1, Number(ym[3]));
    const d = new Date(s);
    return isNaN(+d) ? new Date() : d;
  };
  const first = parseAny(startDateStr);
  if (n === 1) return formatForInput({ id: "any" }, first);

  let end = new Date(first);
  const steps = n - 1;
  if (frequency === "daily") {
    end.setDate(end.getDate() + steps);
  } else if (frequency === "weekly") {
    end.setDate(end.getDate() + steps * 7);
  } else if (frequency === "monthly") {
    const base = new Date(first);
    base.setMonth(base.getMonth() + steps);
    end = base;
  } else {
    return startDateStr;
  }
  return formatForInput({ id: "any" }, new Date(end));
};

const generateSimpleInterestSchedule = (totalRepayable, n) => {
  const schedule = [];
  let cumulativeAmount = 0;

  const standardInstallment = Math.round((totalRepayable / n) * 100) / 100;

  for (let i = 1; i <= n; i++) {
    let installmentAmount;
    if (i < n) {
      installmentAmount = standardInstallment;
      cumulativeAmount += installmentAmount;
    } else {
      installmentAmount = totalRepayable - cumulativeAmount;
    }

    schedule.push({
      installment: i,
      amountDue: +installmentAmount.toFixed(2),
      amountPaid: 0,
      pendingAmount: +installmentAmount.toFixed(2),
      status: "Due",
      paidDate: null,
      modeOfPayment: null,
    });
  }
  return schedule;
};

const generateScheduleWithInstallmentAmount = (
  totalRepayable,
  perInstallment
) => {
  const schedule = [];
  if (!perInstallment || perInstallment <= 0) return schedule;
  const totalCents = Math.round(totalRepayable * 100);
  const perCents = Math.round(perInstallment * 100);
  if (perCents <= 0) return schedule;
  const n = Math.max(1, Math.ceil(totalCents / perCents));
  let remainingCents = totalCents;
  for (let i = 1; i <= n; i++) {
    let amtCents = i < n ? perCents : remainingCents;
    if (i < n) remainingCents -= perCents;
    if (remainingCents < 0) remainingCents = 0;
    const amt = +(amtCents / 100).toFixed(2);
    schedule.push({
      installment: i,
      amountDue: amt,
      amountPaid: 0,
      pendingAmount: amt,
      status: "Due",
      paidDate: null,
      modeOfPayment: null,
    });
  }
  return schedule;
};
