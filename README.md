# Automated Sales Dashboard & Email Report

This project generates a sales performance dashboard from the **Sample Superstore dataset** and automatically emails the results as a html attachment.  

## 📊 Features
- Reads sales data from an Excel file
- Computes key KPIs:
  - Total Sales
  - Total Profit
  - Number of Orders
  - Growth Rate
- Visualizes data using bar charts, maps and line charts
- Exports the dashboard as a **html report**
- Sends the report via **email** (with attachment)

---



## 🚀 How to Run
1. Clone this repository:
   
   git clone https://github.com/NazaarKhalid/Automated_Dashboard.git
   
   cd Automated_Dashboard

   
3. Install dependencies:

   pip install -r requirements.txt


3. Place your dataset in the project folder (e.g., Sample - Superstore Sales (Excel).xlsx).

4. Update the placeholders in the script:

   sender_email = "your_email@example.com"

   sender_password = "your_app_password"

   receiver_email = "receiver_email@example.com"


5. Run the script:

   python main.py

🔒 Security Note

  The email and password in the script are placeholders.
  
  This project is for educational purposes only.
  
  The dataset used (Sample Superstore Sales) is publicly available.
