# 🚀 Audit.Bot_IP: Transforming Clinical Compliance with Zero-Spend Automation

### **Project Goal**

To eliminate a costly, manual compliance audit process and replace it with a **one-click, zero-spend** automation solution built entirely on ubiquitous organizational tools (Excel VBA & Outlook). The objective was to transform a subjective, labor-intensive sampling method into an objective, data-driven system that is **shareable and scalable** across the enterprise.

---

### **The Problem: A Costly, Ineffective Audit**

For years, the Inpatient Pain Audit was a painfully manual exercise that consumed **1-2 hours per day** of expensive Registered Nurse Manager/Charge time.

* **High Labor Cost & Tension:** RNs were forced into a tedious "stare-and-compare" exercise on sampled charts. This process created tension when high-performing staff received low "grades" simply due to **bad luck** in the small, non-representative sampling.
* **No Denominator:** The sampling method provided no denominator information considered whatsoever, meaning the only information generated was a sampling-rate which is ultimately luck of the draw.
* **Data Gaps:** The raw data lives in multiple tables that are not completely mapped to Business Objects, so valuable field data is out of bounds.

---

### **The Solution: Intelligent Automation for Maximum ROI**

The `Audit.Bot_IP` script was engineered to solve a **"thorny problem"** of complex, conditional logic by automating hundreds of processing steps and transforms, delivering several key organizational wins:

#### 1. Zero New Software Spend, Maximum ROI
The entire solution leverages **existing Microsoft Office licenses** (Excel VBA and Outlook), proving that **ubiquitous tools** can be up-purposed to solve complex business problems with **zero new software expenditure**.

#### 2. Comprehensive, Objective Audit
The process eliminates the uncertainty of sampling by handling all available records, providing a complete, objective **denominator** for compliance metrics. This allows management to focus on systemic issues rather than individual staff performance based on chance.

#### 3. Enterprise-Ready Distribution & User Onboarding
The solution was designed to be **clone-to-own** by any analyst at any site:

* **Simple Admin Setup:** A site admin only needs to download the folder structure and run an install script to verify their **individual root path**.
* **Secure User Onboarding (VBA "2FA"):** Invited users confirmed their corporate email address by running a user-install script embedded in their invite file. This automatically:
    * Set up their standardized folder structure.
    * Populated a hidden form with their path information.
    * Returned this information to the admin **via Outlook (behind the scenes)**—a "poor-man's Excel/VBA 2FA" for path verification.
* **User Experience:** Users received a desktop shortcut to a **Link Manager** file, which provided an organized, one-click directory of all program files received, sorted by status (in progress, vs. approved-returned).

---

### **Core Technical Achievements**

The **[Audit\_Logic.vba](https://github.com/kirk-allen-ryan/Clinical_Audit_Automation_VBA/blob/main/Audit_Logic.vba)** script demonstrates expertise in advanced Excel manipulation and algorithmic problem-solving:

* **Algorithmic Scale & Conditional Logic:** The script automates a manual process that involves hundreds of processing steps and transforms. This is powered by approximately **60 unique, multi-level formulas** (including IRF, PRF, and SRF flags) which collectively evaluate over **440,000 gross conditional logic arguments** per audit run. This volume of conditional calculation (IF/AND/OR/VLOOKUP sequencing) is necessary to ensure every event is correctly classified and flagged against 18 separate clinical policy rules.
* **Dynamic Sequencing Logic:** Implements thousands of arguments in logical formulas to sequence events, where a single record can simultaneously serve as a *Pre-PRN assessment* for one event and a *Follow-up assessment* for the previous Medication event.
* **Complex Conditional Rules:** The quantitative value of the pain score itself controls the prerequisites for assessment completeness (Partial vs. Complete).
* **Process Automation:** Automates the voluminous calculations, flag application, and distribution of flagged records back to a trusted user for manual review and final disposition.
