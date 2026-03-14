# Automated Roster Management Tool

A **Flask-based web application** to manage team rosters, leave requests, holiday calendars, and automated notifications. Developed during an internship to streamline workforce management and shift planning.

---

## 🛠 Features

### Employee
- View monthly/yearly shift schedules  
- Apply for planned or unplanned leaves  
- Track personal leave history and approval status  

### Lead
- Upload and preview monthly rosters via Excel (.xlsx)  
- Manage team assignments and function mapping  
- Approve or reject leave requests from their team  
- Upload regional holiday calendars  
- Distribute shift plans via automated email notifications  

### Manager
- Create and organize organizational teams via structured Excel uploads  
- Update roles (promote employees to Leads)  
- Re-assign or remove users from teams  
- Approve or restrict top-level leave requests  
- Generate analytics and download reports for monthly resource allocations  

### Shared Features
- **Email Notifications:** Automatically sends notifications for roster updates, leave approvals, and schedules  
- **Roster Analysis:** Analyze team workloads, shift distribution, and compliance with holidays  

---

## 💻 Tech Stack
- **Backend:** Python, Flask, SQLAlchemy  
- **Database:** MySQL (via SQLAlchemy ORM)  
- **Frontend:** HTML, CSS, Javascript, Bootstrap  
- **Email:** Flask-Mail, SMTP (sample setup)
