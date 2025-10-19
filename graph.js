/*
Copyright (c) 2025 Avdesh Jadon (LoanManager)
All Rights Reserved.
Proprietary and Confidential – Unauthorized copying, modification, or distribution of this file,
via any medium, is strictly prohibited without prior written consent from Avdesh Jadon.
*/

let portfolioChartInstance, profitChartInstance;
Chart.register(ChartDataLabels);
const getThemeColors = () => {
  const computedStyles = getComputedStyle(document.documentElement);

  const parseRgb = (rgbString) => {
    return rgbString.match(/\d+/g).map(Number);
  };

  const hexToRgba = (hex, alpha = 1) => {
    if (!hex || !hex.startsWith("#")) return "rgba(0,0,0,0)";
    let c;
    if (/^#([A-Fa-f0-9]{3}){1,2}$/.test(hex)) {
      c = hex.substring(1).split("");
      if (c.length === 3) {
        c = [c[0], c[0], c[1], c[1], c[2], c[2]];
      }
      c = "0x" + c.join("");
      return `rgba(${[(c >> 16) & 255, (c >> 8) & 255, c & 255].join(
        ","
      )},${alpha})`;
    }
    return "rgba(0,0,0,0)";
  };

  const primaryColor = computedStyles.getPropertyValue("--primary").trim();
  let primaryTransColor = computedStyles
    .getPropertyValue("--primary-t-08")
    .trim();

  if (!primaryTransColor || primaryTransColor.length === 0) {
    if (primaryColor.startsWith("#")) {
      primaryTransColor = hexToRgba(primaryColor, 0.2);
    } else if (primaryColor.startsWith("rgb")) {
      const rgbValues = parseRgb(primaryColor);
      primaryTransColor = `rgba(${rgbValues[0]}, ${rgbValues[1]}, ${rgbValues[2]}, 0.2)`;
    }
  }

  return {
    primary: primaryColor,
    primary_trans: primaryTransColor,
    success: computedStyles.getPropertyValue("--success").trim(),
    danger: computedStyles.getPropertyValue("--danger").trim(),
    grid: computedStyles.getPropertyValue("--border-color").trim(),
    text: computedStyles.getPropertyValue("--text-dark").trim(),
    background: computedStyles.getPropertyValue("--bg-card").trim(),
  };
};

function renderDashboardCharts(activeLoans, settledLoans, profitData) {
  renderPortfolioChart(activeLoans, settledLoans);
  renderProfitChart(profitData);
}

function renderPortfolioChart(activeLoans, settledLoans) {
  const ctx = document.getElementById("portfolioChart")?.getContext("2d");
  if (!ctx) return;

  let totalPaidPrincipal = 0;
  let totalInterestEarned = 0;
  let totalOutstanding = 0;

  [...activeLoans, ...settledLoans].filter(c => c.loanDetails && c.paymentSchedule).forEach(loan => {
      const totalPaid = loan.paymentSchedule.reduce((sum, p) => sum + p.amountPaid, 0);
      const interestPaidOnThisLoan = Math.max(0, totalPaid - loan.loanDetails.principal);
      const principalPaidOnThisLoan = totalPaid - interestPaidOnThisLoan;

      totalInterestEarned += interestPaidOnThisLoan;
      totalPaidPrincipal += principalPaidOnThisLoan;
  });

  activeLoans.filter(c => c.loanDetails && c.paymentSchedule).forEach(loan => {
      const principal = Number(loan.loanDetails.principal || 0);
      const rate = Number(loan.loanDetails.interestRate || 0);
      const given = loan.loanDetails.loanGivenDate;
      const endCandidate = loan.loanDetails.loanEndDate;
      const today = new Date();
      const parsedEnd = (typeof parseDateFlexible === 'function') ? parseDateFlexible(endCandidate) : new Date(endCandidate);
      const endBound = (today < parsedEnd) ? today : parsedEnd;
      const interest = (typeof calculateTotalInterest === 'function')
        ? calculateTotalInterest(principal, rate, given, endBound)
        : 0;
      const totalRepayable = principal + interest;
      const totalPaid = loan.paymentSchedule.reduce((sum, p) => sum + (p.amountPaid || 0), 0);
      totalOutstanding += Math.max(0, totalRepayable - totalPaid);
  });
  
  const theme = getThemeColors();
  const data = {
    labels: ["Outstanding", "Principal Repaid", "Interest Earned"],
    datasets: [
      {
        data: [totalOutstanding, totalPaidPrincipal, totalInterestEarned],
        backgroundColor: [theme.danger, theme.primary, theme.success],
        borderColor: theme.background,
        borderWidth: 4,
        hoverOffset: 16,
        hoverBorderColor: theme.background,
      },
    ],
  };

  if (portfolioChartInstance) portfolioChartInstance.destroy();
  portfolioChartInstance = new Chart(ctx, {
    type: "doughnut",
    data,
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "70%",
      animation: {
        animateScale: true,
        animateRotate: true,
        duration: 1500,
        easing: "easeInOutQuart",
      },
      plugins: {
        legend: {
          position: "bottom",
          labels: { padding: 20, color: theme.text, usePointStyle: true },
        },
        tooltip: {
          callbacks: {
            label: function (context) {
              let label = context.label || "";
              if (label) {
                label += ": ";
              }
              if (context.parsed !== null) {
                label += new Intl.NumberFormat("en-IN", {
                  style: "currency",
                  currency: "INR",
                }).format(context.parsed);
              }
              return label;
            },
          },
        },
        datalabels: { display: false },
      },
    },
  });
}

function renderProfitChart(profitData) {
  const ctx = document.getElementById("profitChart")?.getContext("2d");
  if (!ctx) return;
  const theme = getThemeColors();
  const aggregateData = (data, timeframe) => {
    const aggregated = {};
    data.forEach((item) => {
      const date = new Date(item.date);
      let key;
      if (timeframe === "yearly") {
        key = date.getFullYear().toString();
      } else if (timeframe === "monthly") {
        key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(
          2,
          "0"
        )}`;
      } else {
        key = item.date;
      }
      if (!aggregated[key]) aggregated[key] = 0;
      aggregated[key] += item.profit;
    });
    const sortedKeys = Object.keys(aggregated).sort();
    return {
      labels: sortedKeys,
      values: sortedKeys.map((key) => aggregated[key]),
    };
  };
  const drawChart = (timeframe) => {
    const { labels, values } = aggregateData(profitData, timeframe);
    if (profitChartInstance) {
      profitChartInstance.destroy();
    }
    const gradient = ctx.createLinearGradient(0, 0, 0, ctx.canvas.height);
    gradient.addColorStop(0, theme.primary);
    gradient.addColorStop(1, theme.primary_trans);
    profitChartInstance = new Chart(ctx, {
      type: "line",
      data: {
        labels: labels,
        datasets: [
          {
            label: "Interest Earned",
            data: values,
            backgroundColor: gradient,
            borderColor: theme.primary,
            borderWidth: 2,
            fill: true,
            tension: 0.4,
            pointBackgroundColor: theme.primary,
            pointRadius: 4,
            pointHoverRadius: 6,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          y: {
            beginAtZero: true,
            ticks: { color: theme.text },
            grid: { color: theme.grid },
          },
          x: { ticks: { color: theme.text }, grid: { color: "transparent" } },
        },
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: function (context) {
                return ` Profit: ${new Intl.NumberFormat("en-IN", {
                  style: "currency",
                  currency: "INR",
                }).format(context.parsed.y)}`;
              },
            },
          },
          datalabels: { display: false },
        },
      },
    });
  };
  const controls = document.getElementById("profit-chart-controls");
  if (controls) {
    controls.querySelectorAll("button").forEach((button) => {
      button.addEventListener("click", (e) => {
        if (e.currentTarget.classList.contains("active")) return;
        controls.querySelector("button.active").classList.remove("active");
        e.currentTarget.classList.add("active");
        drawChart(e.currentTarget.dataset.frame);
      });
    });
  }
  drawChart("monthly");
}