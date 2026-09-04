# ⚡ Automated B2B Order & Payment Reconciliation Engine

An automated financial audit pipeline that cross-references internal store orders (Shopify, WooCommerce, ERPs) against external payment gateway settlement ledgers (Stripe, Razorpay, Boleto, etc.). 
readme
It reconciles multi-thousand-row transaction files in **under 2 seconds**, flagging uncollected revenue, gateway charge discrepancies, and unrefunded canceled orders into an audit-ready Excel workbook.

---

## 🛑 The Core Problem
Growing e-commerce and retail brands spend **15–20 hours every month** manually matching order IDs against payment settlements. This manual approach causes:
* **Silent Revenue Leaks:** Orders marked "completed" where gateway settlements failed or dropped.
* **Accounting Liabilities:** Canceled/out-of-stock orders that were never refunded, triggering chargebacks and fines.
* **Pricing Mismatches:** Surcharges, partial payments, or currency conversions that do not match the internal order value.

---

## 📊 Business Proof & Audit Output

### 1. Executive Summary & Exposure Calculation
Instant visibility into total reconciled volume vs. at-risk capital exposure:
![Executive Summary](assets/executive_summary.png)

### 2. High-Volume Audit Trail (Color-Coded Flags)
Automatically highlights mismatches (yellow) and missing/orphaned settlements (peach/red):
![Audit Trail](assets/audit_trail_discrepancy_1.png)

---

## ⚙️ Core Technical Capabilities
* **Multi-Rail Aggregation:** Combines multiple vouchers, cards, and split payments per single order ID.
* **Variance & Status Classification:** Categorizes records into `SETTLED_MATCH`, `AMOUNT_MISMATCH`, `MISSING_IN_GATEWAY`, `CANCELED_BUT_CHARGED`, and `ORPHANED_PAYMENT`.
* **Automated Excel Formatting:** Generates executive KPI tabs and fully styled audit worksheets with auto-fitted column widths using `openpyxl`.

---

## 🚀 Quickstart

```bash
# Clone the repository
git clone [https://github.com/INDRONIIL/b2b-ecommerce-financial-reconciliation.git](https://github.com/INDRONIIL/b2b-ecommerce-financial-reconciliation.git)

# Navigate to project directory
cd b2b-ecommerce-financial-reconciliation

# Install dependencies
pip install -r requirements.txt

# Run the reconciliation engine
python reconcile.py