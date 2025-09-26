import pandas as pd 
import matplotlib.pyplot as plt 
import plotly.express as px 
import base64 
from io import BytesIO 
from IPython.display import display, HTML 
import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 
 
 
df = pd.read_excel("Sample - Superstore Sales (Excel).xlsx", 
sheet_name="Orders") 
df.dropna(subset=['Sales', 'Customer Name', 'Product Name'], inplace=True) 
df['Order Date'] = pd.to_datetime(df['Order Date']) 
df['Week'] = df['Order Date'].dt.isocalendar().week 
 
latest_week = df['Week'].max() 
prev_week = latest_week - 1 
 
current_week_df = df[df['Week'] == latest_week] 
previous_week_df = df[df['Week'] == prev_week] 
 
# KPIs 
total_products_sold = current_week_df['Order Quantity'].sum() 
total_sales = current_week_df['Sales'].sum() 
total_profit = current_week_df['Profit'].sum() 
 
prev_week_sales = previous_week_df['Sales'].sum() 
growth_rate = ((total_sales - prev_week_sales) / prev_week_sales) * 100 if 
prev_week_sales > 0 else 0 
 
# --------- Line Chart --------- 
current_week_sales = current_week_df.groupby(current_week_df['Order 
Date'].dt.dayofweek)['Sales'].sum().reindex(range(7), fill_value=0) 
previous_week_sales = previous_week_df.groupby(previous_week_df['Order 
Date'].dt.dayofweek)['Sales'].sum().reindex(range(7), fill_value=0) 
 
plt.figure(figsize=(8,4)) 
plt.plot(current_week_sales.index, current_week_sales.values, marker='o', 
label='Current Week') 
plt.plot(previous_week_sales.index, previous_week_sales.values, 
marker='o', label='Previous Week') 
plt.xticks(range(7), ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']) 
plt.ylabel('Sales ($)') 
plt.title('Weekly Sales Comparison') 
plt.legend() 
plt.grid(True) 
plt.tight_layout() 
 
buf = BytesIO() 
plt.savefig(buf, format='png') 
plt.close() 
buf.seek(0) 
line_chart_base64 = base64.b64encode(buf.read()).decode('utf-8') 
 
# --------- Top 5 Products Table --------- 
top_products = current_week_df.groupby('Product 
Name')['Sales'].sum().sort_values(ascending=False).head(5).reset_index() 
top_products['Sales'] = top_products['Sales'].round(0).astype(int) 
 
top_products_html = ( 
    top_products.style 
    .set_table_styles( 
        [ 
            {'selector': 'th', 
             'props': [('background-color', 'black'), 
                       ('color', 'white'), 
                       ('text-align', 'center'), 
                       ('padding', '8px')]}, 
            {'selector': 'td', 
             'props': [('text-align', 'center'), 
                       ('padding', '8px')]}, 
            {'selector': 'tbody tr:nth-child(odd)', 
             'props': [('background-color', '#d9d9d9'), 
                       ('color', 'black')]}, 
            {'selector': 'tbody tr:nth-child(even)', 
             'props': [('background-color', 'white'), 
                       ('color', 'black')]}, 
        ] 
    ) 
    .hide(axis='index') 
    .to_html() 
) 
 
# --------- Top 3 Customers Table --------- 
top_customers = current_week_df.groupby('Customer 
Name')['Sales'].sum().sort_values(ascending=False).head(5).reset_index() 
top_customers['Sales'] = top_customers['Sales'].round(0).astype(int) 
 
top_customers_html = ( 
    top_customers.style 
    .set_table_styles( 
        [ 
            {'selector': 'th', 
             'props': [('background-color', 'black'), 
                       ('color', 'white'), 
                       ('text-align', 'center'), 
                       ('padding', '8px')]}, 
            {'selector': 'td', 
             'props': [('text-align', 'center'), 
                       ('padding', '8px')]}, 
            {'selector': 'tbody tr:nth-child(odd)', 
             'props': [('background-color', '#d9d9d9'), 
                       ('color', 'black')]}, 
            {'selector': 'tbody tr:nth-child(even)', 
             'props': [('background-color', 'white'), 
                       ('color', 'black')]}, 
        ] 
    ) 
    .hide(axis='index') 
    .to_html() 
) 
 
# --------- Map (plotly) --------- 
import plotly.io as pio 
 
state_name_to_code = { 
    'Alabama': 'AL', 'Alaska': 'AK', 'Arizona': 'AZ', 'Arkansas': 'AR', 
    'California': 'CA', 'Colorado': 'CO', 'Connecticut': 'CT', 'Delaware': 
'DE', 
    'Florida': 'FL', 'Georgia': 'GA', 'Hawaii': 'HI', 'Idaho': 'ID', 
    'Illinois': 'IL', 'Indiana': 'IN', 'Iowa': 'IA', 'Kansas': 'KS', 
    'Kentucky': 'KY', 'Louisiana': 'LA', 'Maine': 'ME', 'Maryland': 'MD', 
    'Massachusetts': 'MA', 'Michigan': 'MI', 'Minnesota': 'MN', 
'Mississippi': 'MS', 
    'Missouri': 'MO', 'Montana': 'MT', 'Nebraska': 'NE', 'Nevada': 'NV', 
    'New Hampshire': 'NH', 'New Jersey': 'NJ', 'New Mexico': 'NM', 'New 
York': 'NY', 
    'North Carolina': 'NC', 'North Dakota': 'ND', 'Ohio': 'OH', 
'Oklahoma': 'OK', 
    'Oregon': 'OR', 'Pennsylvania': 'PA', 'Rhode Island': 'RI', 
    'South Carolina': 'SC', 'South Dakota': 'SD', 'Tennessee': 'TN', 
    'Texas': 'TX', 'Utah': 'UT', 'Vermont': 'VT', 'Virginia': 'VA', 
    'Washington': 'WA', 'West Virginia': 'WV', 'Wisconsin': 'WI', 
'Wyoming': 'WY' 
} 
 
df['State_Code'] = df['State'].map(state_name_to_code).fillna(df['State']) 
 
state_sales = ( 
    df.groupby('State_Code')['Sales'] 
    .sum() 
    .reset_index() 
) 
 
fig_state_sales = px.choropleth( 
    state_sales, 
    locations='State_Code', 
    locationmode='USA-states', 
    color='Sales', 
    scope='usa', 
    color_continuous_scale='Greys', 
    labels={'Sales': 'Total Sales'}, 
    title='Sales by State' 
) 
 
fig_state_sales.update_traces(marker_line_color='black', 
marker_line_width=0.8) 
 
fig_state_sales.update_layout( 
    autosize=True, 
    margin=dict(l=20, r=20, t=50, b=20), 
) 
 
 
map_html = pio.to_html( 
    fig_state_sales, 
    full_html=False, 
    include_plotlyjs='cdn', 
    config={'responsive': True} 
) 
 
# --------- Final dashboard HTML --------- 
dashboard_html = f""" 
<html> 
<head> 
  <title>Sales Dashboard</title> 
  <style> 
    body {{ font-family: Arial, sans-serif; margin: 20px; }} 
    .dashboard {{ max-width: 1200px; border: 2px solid black; padding: 
20px; }} 
    h1 {{ text-align: center; margin-bottom: 30px; }} 
    .kpis {{ display: flex; justify-content: space-around; margin-bottom: 
30px; }} 
    .kpi {{ border: 1px solid #ccc; padding: 15px 25px; text-align: 
center; width: 22%; background: #f9f9f9; font-size: 1.2em; }} 
    .charts-top {{ display: flex; justify-content: space-between; gap: 
20px; margin-bottom: 30px; }} 
    .chart {{ border: 1px solid #ccc; padding: 10px; width: 48%; }} 
    .tables-bottom {{ display: flex; justify-content: space-between; gap: 
20px; margin-bottom: 30px; }} 
 
    /* Make Sales by State container wider, Customers narrower */ 
    .tables-bottom .table-container:first-child {{ 
      border: 1px solid #ccc; 
      padding: 10px; 
      width: 65%; 
      box-sizing: border-box; 
      position: relative; 
      overflow: hidden; 
      height: 620px; /* slightly bigger than plot height */ 
    }} 
 
    .tables-bottom .table-container:first-child .plotly-graph-div {{ 
      width: 100% !important; 
      height: 600px !important; 
      box-sizing: border-box; 
      max-width: 100%; 
      max-height: 100%; 
    }} 
 
    .tables-bottom .table-container:last-child {{ 
      border: 1px solid #ccc; 
      padding: 10px; 
      width: 33%; 
      box-sizing: border-box; 
    }} 
  </style> 
 
</head> 
<body> 
  <div class="dashboard"> 
    <h1>Weekly Sales Analysis</h1> 
 
    <div class="kpis"> 
      <div class="kpi"><strong>{total_products_sold}</strong><br>Total 
Products Sold </div> 
      <div class="kpi"><strong>${total_sales:,.0f}</strong><br>Total 
Sales</div> 
      <div class="kpi"><strong>${total_profit:,.0f}</strong><br>Total 
Profit</div> 
      <div class="kpi"><strong>{growth_rate:.2f}%</strong><br>Growth 
Rate</div> 
    </div> 
 
    <div class="charts-top"> 
      <div class="chart"> 
        <h3>Weekly Sales Comparison</h3> 
        <img src="data:image/png;base64,{line_chart_base64}" 
style="width:100%;" /> 
      </div> 
 
      <div class="chart"> 
        <h3>Top 5 Products by Sales</h3> 
        {top_products_html} 
      </div> 
    </div> 
 
    <div class="tables-bottom"> 
      <div class="table-container"> 
        <h3>Sales by State</h3> 
        {map_html} 
      </div> 
      <div class="table-container"> 
        <h3>Top 5 Customers by Sales</h3> 
        {top_customers_html} 
      </div> 
    </div> 
  </div> 
</body> 
</html> 
""" 
 
with open("dashboard.html", "w") as f: 
    f.write(dashboard_html) 
 
 

smtp_server = 'smtp.gmail.com'  
smtp_port = 587 
sender_email = "your_email@example.com"
sender_password = "your_app_password"
receiver_email = "receiver_email@example.com" 
subject = "Weekly Sales Dashboard" 
 
msg = MIMEMultipart() 
msg['From'] = sender_email 
msg['To'] = receiver_email 
msg['Subject'] = subject 
 
body = "Hi,\n\nPlease find the weekly sales dashboard attached.\n\nBest 
Nazaar Khalid." 
msg.attach(MIMEText(body, 'plain')) 
 
filename = "dashboard.html" 
with open(filename, "rb") as attachment: 
    part = MIMEBase('application', 'octet-stream') 
    part.set_payload(attachment.read()) 
 
encoders.encode_base64(part) 
 
part.add_header( 
    "Content-Disposition", 
    f"attachment; filename= {filename}", 
) 
 
msg.attach(part) 
 
try: 
    server = smtplib.SMTP(smtp_server, smtp_port) 
    server.starttls() 
    server.login(sender_email, sender_password) 
    server.send_message(msg) 
    print("Email sent successfully!") 
except Exception as e: 
    print(f"Error sending email: {e}") 
finally: 
    server.quit()
