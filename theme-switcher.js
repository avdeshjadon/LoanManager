/*
Copyright (c) 2025 Avdesh Jadon (LoanManager)
All Rights Reserved.
Proprietary and Confidential – Unauthorized copying, modification, or distribution of this file,
via any medium, is strictly prohibited without prior written consent from Avdesh Jadon.
*/

document.addEventListener("DOMContentLoaded", () => {
  const changeThemeBtn = document.getElementById("change-theme-btn");
  const themeModal = document.getElementById("theme-modal");
  const lightThemesGrid = document.getElementById("light-themes-grid");
  const darkThemesGrid = document.getElementById("dark-themes-grid");

  if (!changeThemeBtn || !themeModal) return;

  const themes = [
    {
      name: "Default",
      light: {
        id: "default",
        colors: { p: "#4a55a2", s: "#6d28d9", bm: "#f1f5f9", bc: "#ffffff" },
      },
      dark: {
        id: "dark",
        colors: { p: "#818cf8", s: "#a78bfa", bm: "#0f172a", bc: "#1e293b" },
      },
    },
    {
      name: "Midnight Sapphire",
      light: {
        id: "midnight-sapphire",
        colors: { p: "#1e40af", s: "#3730a3", bm: "#f8fafc", bc: "#ffffff" },
      },
      dark: {
        id: "midnight-sapphire-dark",
        colors: { p: "#60a5fa", s: "#818cf8", bm: "#0f172a", bc: "#1e293b" },
      },
    },
    {
      name: "Crimson Blaze",
      light: {
        id: "crimson-blaze",
        colors: { p: "#dc2626", s: "#b91c1c", bm: "#fef2f2", bc: "#ffffff" },
      },
      dark: {
        id: "crimson-blaze-dark",
        colors: { p: "#f87171", s: "#ef4444", bm: "#450a0a", bc: "#7f1d1d" },
      },
    },
    {
      name: "Emerald Oasis",
      light: {
        id: "emerald-oasis",
        colors: { p: "#059669", s: "#047857", bm: "#f0fdf4", bc: "#ffffff" },
      },
      dark: {
        id: "emerald-oasis-dark",
        colors: { p: "#34d399", s: "#10b981", bm: "#052e16", bc: "#064e3b" },
      },
    },
    {
      name: "Royal Amethyst",
      light: {
        id: "royal-amethyst",
        colors: { p: "#7c3aed", s: "#6d28d9", bm: "#faf5ff", bc: "#ffffff" },
      },
      dark: {
        id: "royal-amethyst-dark",
        colors: { p: "#a78bfa", s: "#8b5cf6", bm: "#1e1b4b", bc: "#312e81" },
      },
    },
    {
      name: "Sunset Orange",
      light: {
        id: "sunset-orange",
        colors: { p: "#ea580c", s: "#c2410c", bm: "#fff7ed", bc: "#ffffff" },
      },
      dark: {
        id: "sunset-orange-dark",
        colors: { p: "#fb923c", s: "#f97316", bm: "#431407", bc: "#7c2d12" },
      },
    },
    {
      name: "Ocean Teal",
      light: {
        id: "ocean-teal",
        colors: { p: "#0d9488", s: "#0f766e", bm: "#f0fdfa", bc: "#ffffff" },
      },
      dark: {
        id: "ocean-teal-dark",
        colors: { p: "#2dd4bf", s: "#14b8a6", bm: "#042f2e", bc: "#115e59" },
      },
    },
  ];

  window.applyTheme = function (themeId) {
    document.documentElement.setAttribute("data-theme", themeId);
    localStorage.setItem("theme", themeId);

    document
      .querySelectorAll(".theme-card")
      .forEach((card) => card.classList.remove("active"));
    const activeCard = document.querySelector(
      `.theme-card[data-theme="${themeId}"]`
    );
    if (activeCard) {
      activeCard.classList.add("active");
    }

    const isDarkMode = themeId.endsWith("-dark") || themeId === "dark";

    const headerToggleBtn = document.getElementById("theme-toggle-btn");
    if (headerToggleBtn) {
      headerToggleBtn.innerHTML = isDarkMode
        ? '<i class="fas fa-sun"></i>'
        : '<i class="fas fa-moon"></i>';
    }

    const settingsToggleSwitch = document.getElementById("dark-mode-toggle");
    if (settingsToggleSwitch) {
      settingsToggleSwitch.checked = isDarkMode;
    }

    if (typeof renderDashboardCharts === "function" && window.allCustomers) {
      const profitData = window.processProfitData(window.allCustomers);
      renderDashboardCharts(
        window.allCustomers.active,
        window.allCustomers.settled,
        profitData
      );
    }
  };

  window.toggleDarkMode = function () {
    const currentThemeId = localStorage.getItem("theme") || "default";
    const themeInfo = themes.find(
      (t) => t.light.id === currentThemeId || t.dark.id === currentThemeId
    );

    if (themeInfo) {
      const isCurrentlyDark = currentThemeId === themeInfo.dark.id;
      window.applyTheme(
        isCurrentlyDark ? themeInfo.light.id : themeInfo.dark.id
      );
    }
  };

  function createThemeCard(themeVariant, name) {
    const card = document.createElement("div");
    card.className = "theme-card";
    card.dataset.theme = themeVariant.id;
    card.innerHTML = `
            <div class="theme-name">${name}</div>
            <div class="theme-palette">
                <div class="theme-color-swatch" style="background-color: ${themeVariant.colors.p}"></div>
                <div class="theme-color-swatch" style="background-color: ${themeVariant.colors.s}"></div>
                <div class="theme-color-swatch" style="background-color: ${themeVariant.colors.bm}"></div>
                <div class="theme-color-swatch" style="background-color: ${themeVariant.colors.bc}"></div>
            </div>
        `;
    card.addEventListener("click", () => {
      window.applyTheme(themeVariant.id);
      themeModal.classList.remove("show");
    });
    return card;
  }
  function populateThemeModal() {
    if (!lightThemesGrid || !darkThemesGrid) return;
    lightThemesGrid.innerHTML = "";
    darkThemesGrid.innerHTML = "";

    themes.forEach((theme) => {
      lightThemesGrid.appendChild(createThemeCard(theme.light, theme.name));
      darkThemesGrid.appendChild(
        createThemeCard(theme.dark, `${theme.name} (Dark)`)
      );
    });
  }
  populateThemeModal();
  const savedTheme = localStorage.getItem("theme") || "default";
  window.applyTheme(savedTheme);

  changeThemeBtn.addEventListener("click", (e) => {
    e.preventDefault();
    themeModal.classList.add("show");
  });
});
