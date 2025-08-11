# 💉 Hospital Management System - Microsoft Access Project(VBA)

A simple yet fully functional **Hospital Management System** built using **Microsoft Access**, featuring relational database design, user-friendly forms, VBA automation, queries, and reports. This project showcases end-to-end patient management, doctor assignment, billing, and appointment scheduling.

---

## 📂 Project Structure

### 🧾 Tables

| Table Name     | Description                                                  |
|----------------|--------------------------------------------------------------|
| `Patients`     | Stores patient information (Full Name, Date of Birth, Gender) |
| `Doctors`      | Contains doctor details (Name, Specialization, Contact)      |
| `Appointments` | Links patients and doctors with appointment date & time      |
| `Billing`      | Tracks total amount, paid amount, and outstanding balance    |

---

### 🧮 Queries

| Query Name                | Purpose                                              |
|---------------------------|------------------------------------------------------|
| `UpcomingAppointments`    | Lists all appointments scheduled in the future       |
| `UnpaidBills`             | Shows patients with pending bills                    |
| `DoctorSchedule`          | Displays appointments for a specific doctor          |

---

### 📝 Forms

| Form Name           | Functionality                                               |
|---------------------|-------------------------------------------------------------|
| `PatientEntryForm`  | Input form for registering new patients with validation     |
| `DoctorEntryForm`   | Form for adding or updating doctor information              |
| `AppointmentForm`   | Book appointments via dropdowns (combo boxes) with logic    |
| `BillingForm`       | Enter billing details; shows thank you or reminder alerts  |

---

### 📊 Reports

| Report Name             | Description                                               |
|--------------------------|-----------------------------------------------------------|
| `PatientListReport`      | Printable list of all registered patients                 |
| `AppointmentSchedule`    | Summary of all upcoming appointments                      |
| `BillingSummaryReport`   | Report of all bills and payment status                    |

---

## ⚙️ VBA Functionalities Used

| Concept                  | Example Used                                              |
|--------------------------|-----------------------------------------------------------|
| `Form Load Message`      | Displays welcome message on opening form                  |
| `Validation Checks`      | Ensures required fields are not empty                     |
| `Message Boxes`          | Shows user feedback (success, warning)                    |
| `Conditional Logic`      | Shows "Thank You" if bill is paid; else prompts alert     |
| `Combo Box Bound Column` | Stores ID instead of names to ensure database normalization |
| `Form Events`            | Uses `BeforeUpdate`, `OnClick`, and `OnLoad` events       |
| `Custom Functions`       | Age calculation from DOB using VBA                        |

---

## ✅ Example VBA Snippets

```vba
' Display welcome message on form load
Private Sub Form_Load()
    MsgBox "Welcome to the Hospital Management System!"
End Sub
```
---

## 💡 Future Enhancements

- Enable automated SMS/email alerts for upcoming appointments and billing due dates.
- Implement role-based access control for Admin, Doctors, and Receptionists.
- Integrate biometric or QR code patient check-in for enhanced workflow.
- Expand reporting module to include doctor-specific and department-level analytics.
- Embed Power BI dashboards into Access or link to external analytics system.

---

## 🤝 Let's Connect

I'm always open to feedback, mentorship, and collaboration opportunities!

- 🔗 [LinkedIn – K Jaya Krishna] (www.linkedin.com/in/k-jaya-krishna-b675a2229)
- 💻 [GitHub – JayaKrishna003](https://github.com/JayaKrishna003)

---
## 🔖 Tags

#HospitalManagementSystem #MicrosoftAccess #VBA #DatabaseDesign #HealthTech #PatientRecords #AccessForms #AccessReports #VBAValidation #BillingAutomation #MedicalDatabase #JayaKrishnaProjects #MSAccessProjects


