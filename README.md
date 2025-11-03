# Loan Manager: Advanced EMI Loan Management System

[![Live Demo](https://img.shields.io/badge/Live_Demo-Launch_Application-00C853?style=for-the-badge&logo=vercel)](https://loan-manager-main.vercel.app/)
[![License](https://img.shields.io/badge/License-Proprietary-E91E63?style=for-the-badge)](./LICENSE)
[![Technology](https://img.shields.io/badge/Made_with-JavaScript_&_Firebase-F7DF1E?style=for-the-badge&logo=javascript)](https://firebase.google.com/)

**Loan Manager** is a comprehensive, single-page web application designed for the meticulous management of Equated Monthly Installment (EMI) based loans. Built with Vanilla JavaScript and powered by Google's Firebase suite, this application provides a secure, and user-friendly platform for individual moneylenders and small-scale financial entities to manage their lending operations with precision and ease.

The application moves beyond simple record-keeping, offering an interactive dashboard, detailed customer and loan analytics, full data portability, and a secure, authenticated environment for data integrity.

*(**Note:** This is a sample screenshot. You can replace the link above with a real screenshot of your application's dashboard for a better visual representation.)*

---

## 🚀 Live Application

Experience the full functionality of the application hosted on Vercel:
**[https://loan-manager-main.vercel.app/](https://loan-manager-nine.vercel.app/)**

---

## ✨ Core Features & Functionality

This application is packed with features designed to provide a complete loan management experience.

### 1. **Secure Authentication System**
   - **Firebase-Powered:** Utilizes Firebase Authentication for a secure and reliable email and password login system.
   - **Session Persistence:** Users remain logged in across sessions for a seamless experience.
   - **Password Management:** Includes a "Forgot Password" functionality for secure account recovery.

### 2. **Interactive & Insightful Dashboard**
   - **Key Performance Indicators (KPIs):** At-a-glance cards display critical metrics like *Total Principal Disbursed*, *Total Outstanding Balance*, *Total Interest Earned*, and *Active Loan Count*.
   - **Portfolio Visualization:** A dynamic doughnut chart visualizes the composition of your entire loan portfolio (e.g., Outstanding vs. Repaid Principal vs. Interest).
   - **Profitability Analysis:** An elegant line chart tracks interest earned over time, with options to view data aggregated by month or year.

### 3. **Comprehensive Customer & Loan Management**
   - **Centralized Customer Database:** Manage all customers from a single interface, with a clear separation between `Active` and `Settled` accounts.
   - **Dynamic Search:** Instantly find any active customer by name.
   - **Detailed Customer View:** A dedicated modal provides a 360-degree view of each customer, including personal and KYC details, a loan progress bar, a complete EMI amortization schedule, and financial summaries.

### 4. **Granular EMI Tracking**
   - **Full Amortization Schedule:** Automatically generates a detailed EMI schedule upon loan creation, breaking down each payment into principal and interest components.
   - **One-Click Status Updates:** Mark EMIs as `Paid` or `Unpaid`. The system includes validation to prevent out-of-sequence payments.
   - **Automatic Status Highlighting:** The UI automatically flags `Paid`, `Due`, and `Overdue` EMIs with distinct colors for immediate visual feedback.

### 5. **Advanced Loan Operations**
   - **Loan Refinancing:** Seamlessly refinance an existing loan. The system calculates the outstanding balance, allows for additional principal to be added, and generates a new loan schedule under new terms.
   - **Refinance History:** All previous loan iterations are saved, providing a complete historical trail for refinanced accounts.
   - **Loan Settlement:** Formally close a loan by moving it from the "Active" to the "Settled" list, which locks the EMI schedule from further changes.

### 6. **Robust Data Management & Portability**
   - **Export to Excel:** Export detailed reports of all active or settled customers into a formatted `.xlsx` file with a single click.
   - **Full Data Backup:** Generate a complete JSON backup of all application data for safekeeping.
   - **Data Restoration:** Restore the application to a previous state by importing a JSON backup file. This feature includes multiple confirmations to prevent accidental data overwrites.

### 7. **User Experience & Utilities**
   - **Integrated EMI Calculator:** A handy tool to quickly calculate the EMI, total interest, and total payment for any loan amount, rate, and tenure.
   - **Dual-Theme Interface:** Toggle between a sleek **Dark Mode** and a clean **Light Mode** to suit your preference.
   - **Real-time Notifications:** A non-intrusive toast notification system provides feedback for all major actions.
   - **Fully Responsive Design:** Meticulously designed to provide a consistent and intuitive experience across all devices, from large desktops to small mobile phones.

---

## 🛠️ Technology & Architecture

The project is built on a foundation of modern, efficient, and scalable web technologies.

-   **Front-End:**
    -   `HTML5`: For the core structure and content.
    -   `CSS3`: For custom styling, animations, and responsive design.
    -   `Vanilla JavaScript (ES6+)`: For all client-side logic, DOM manipulation, and application state management.

-   **Back-End & Database:**
    -   **Firebase Firestore:** A NoSQL cloud database used for storing all customer and loan data in real-time.
    -   **Firebase Authentication:** Handles user registration, login, and session management.

-   **Libraries & Utilities:**
    -   **Chart.js:** For creating beautiful and interactive charts on the dashboard.
    -   **SheetJS (xlsx.js):** A powerful library for generating Excel files directly in the browser.
    -   **Font Awesome:** For a rich icon set used throughout the user interface.

---

## ⚙️ Getting Started: Local Development Setup

To get a local copy up and running, follow these simple steps.

### Prerequisites

-   A modern web browser (e.g., Google Chrome, Firefox).
-   A Google Firebase account (the free "Spark Plan" is sufficient).

### Installation & Configuration

1.  **Clone the Repository**
    ```sh
    git clone [https://github.com/avdeshjadon/LoanManager.git](https://github.com/avdeshjadon/LoanManager.git)
    ```

2.  **Navigate to the Project Directory**
    ```sh
    cd LoanManager
    ```

3.  **Configure Firebase**
    This is the most critical step. The application needs to connect to your own Firebase instance.
    -   Go to the [Firebase Console](https://console.firebase.google.com/) and create a new project.
    -   Inside your new project, add a new **Web App**.
    -   Firebase will provide you with a `firebaseConfig` object. Copy this entire object.
    -   Open the `script.js` file in your code editor.
    -   Replace the placeholder `firebaseConfig` object at the top of the file with the one you copied from your Firebase project.
    ```javascript
    const firebaseConfig = {
      apiKey: "YOUR_API_KEY",
      authDomain: "YOUR_AUTH_DOMAIN",
      projectId: "YOUR_PROJECT_ID",
      storageBucket: "YOUR_STORAGE_BUCKET",
      messagingSenderId: "YOUR_MESSAGING_SENDER_ID",
      appId: "YOUR_APP_ID"
    };
    ```
    -   In the Firebase Console, navigate to **Authentication** -> **Sign-in method** and enable the **Email/Password** provider.
    -   Navigate to **Firestore Database**, create a new database, and start it in **Production mode**.

4.  **Run the Application**
    -   No complex build steps are required. Simply open the `index.html` file in your web browser.

---

## 📂 Project File Structure

```
├── index.html            # Main HTML file, the entry point of the application
├── style.css             # All CSS styles for layout, theme, and responsiveness
├── script.js             # Core application logic, event listeners, Firebase integration
├── graph.js              # Logic for rendering dashboard charts using Chart.js
├── excel.js              # Contains the function for exporting data to Excel
├── LICENSE               # The proprietary license for the project
└── README.md             # This file
```


---

## 📄 License

This project is protected under a **Proprietary License**.

© 2025 Avdesh Jadon (Loanmanager)

All rights are reserved. You may use the live version of this software for personal, non-commercial purposes. However, you are **strictly prohibited** from copying, modifying, distributing, or using this source code for any commercial purpose.

For more details, please see the `LICENSE` file.

---

Built with ❤️ by **Avdesh Jadon** and ***Sachin Yadav**
