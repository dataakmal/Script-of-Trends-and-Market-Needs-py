#!/usr/bin/env python
# coding: utf-8

# In[1]:


#Import Library
from datetime import datetime,date,timedelta
import pandas as pd
from datetime import datetime
import os
import numpy as np  # Import NumPy
import matplotlib.pyplot as plt
import seaborn as sns


# In[2]:


# Source Data
root = r'C:\Users\think\Downloads'
dirs = [file for file in os.listdir(root) if 'Recruitment Trends and Market Demand Analysis' in file and '~$' not in file]

# Membaca data dari setiap file
demand, candidate, position = None, None, None

for xlsb in dirs:
    file = os.path.join(root, xlsb)
    
    # Membaca sheet yang diperlukan
    demand = pd.read_excel(file, sheet_name='monthly_demand_data', usecols="A:C")
    candidate = pd.read_excel(file, sheet_name='candidate_data', usecols="A:F")
    position = pd.read_excel(file, sheet_name='position_data', usecols="A:F")


# In[3]:


#Cleansing and Formatting Data
# Mengonversi kolom date menjadi tipe datetime
demand['date'] = pd.to_datetime(demand['date'])
demand['year'] = demand['date'].dt.strftime('%Y')

#Menghapus kolom NaN
demand = demand.dropna(subset=['industry'])
demand = demand.dropna(subset=['demand_index'])
demand['year'] = pd.to_numeric(demand['year'], errors='coerce')
demand['industry']     = demand['industry'].str.upper()
demand['demand_index'] = demand['demand_index'].astype(int)


# In[4]:


# Create industry_trend DataFrame
industry_trend = demand.pivot_table(
    index='year',
    columns='industry',
    values='demand_index',
    aggfunc='median')

plt.figure(figsize=(12, 6))
palette = sns.color_palette("husl", len(industry_trend.columns))

# Plot data
ax = industry_trend.plot(ax=plt.gca(), color=palette, linewidth=2)

# Add labels at data points below a threshold
for industry in industry_trend.columns:
    for year, value in industry_trend[industry].dropna().iteritems():
        if value <= 1000:  # Add labels only if the value is below a Y-axis threshold
            plt.text(int(year), value + 20, f'{value:.0f}', 
                     color='black', fontsize=9, ha='center')

# Ensure x-axis displays only integer years
plt.xticks(ticks=industry_trend.index, labels=[str(int(year)) for year in industry_trend.index])
plt.title('Trends of Demand by Industry', fontsize=16, fontweight='bold')
plt.xlabel('Year', fontsize=12)
plt.ylabel('Median Demand Index', fontsize=12)
plt.legend(title='Industry', loc='upper left', fontsize=10)
plt.grid(color='gray', linestyle='--', linewidth=0.5, alpha=0.7)

# Limit Y-axis to a specific range
plt.ylim(0, 400)
plt.tight_layout()
plt.show()

# Calculate median demand per industry and sort by value
industry_demand = demand.groupby('industry')['demand_index'].median().sort_values(ascending=False)

# Plot top industries by median demand index
plt.figure(figsize=(8, 5))
sns.barplot(x=industry_demand.values, y=industry_demand.index, palette="viridis")

# Add title and axis labels
plt.title('Top Industries by Median Demand Index', fontsize=16, fontweight='bold')
plt.xlabel('Median Demand Index', fontsize=12)
plt.ylabel('Industry', fontsize=12)
plt.grid(color='gray', linestyle='--', linewidth=0.5, alpha=0.7)

# Add values to each bar
for index, value in enumerate(industry_demand.values):
    plt.text(value, index, f'{value:.1f}', va='center', ha='left', fontsize=10)

# Adjust layout for better visualization
plt.tight_layout()
plt.show()


# In[5]:


poscandidate = pd.merge(candidate, position, on='position_id', how='left')
poscandidate = poscandidate.drop(columns=['position_id'])

#Menghapus kolom NaN
poscandidate = poscandidate.dropna(subset=['avg_salary'])
poscandidate['industry']         = poscandidate['industry'].str.upper()
poscandidate['job_title']        = poscandidate['job_title'].str.upper()
poscandidate['location']         = poscandidate['location'].str.upper()
poscandidate['experience_level'] = poscandidate['experience_level'].str.upper()
poscandidate['education_level']  = poscandidate['education_level'].str.upper()
poscandidate['avg_salary']       = poscandidate['avg_salary'].astype(int)


# In[6]:


# Menghitung frekuensi keterampilan yang dibutuhkan
job_demand = poscandidate['job_title'].value_counts()
job_demand_percentage = poscandidate['job_title'].value_counts(normalize=True) * 100

# Gabungkan hasil hitung jumlah dan persentase
job_demand_with_percentage = pd.DataFrame({
    'Count': job_demand,
    'Percentage': job_demand_percentage
})


# In[7]:


# Format persentase dengan dua desimal
job_demand_with_percentage['Percentage'] = job_demand_with_percentage['Percentage'].apply(lambda x: f"{x:.2f}%")
# --- Donut Chart ---
plt.figure(figsize=(8, 8))
# Donut chart dengan palet warna 'Set2' dari seaborn
sns.set_palette("Set2")

# Membuat pie chart dengan warna yang lebih modern
plt.pie(job_demand_percentage, labels=job_demand_percentage.index, autopct='%1.1f%%', startangle=140, wedgeprops={'width': 0.3})

# Menambahkan judul dengan style
plt.title('Distribution of Most Frequently Needed Skills', fontsize=18, fontweight='bold')

# Menambahkan garis tengah untuk mengubah pie menjadi donut chart
plt.gca().add_artist(plt.Circle((0, 0), 0.55, color='white'))

# Menampilkan plot
plt.axis('equal')  # Menjaga proporsi pie agar tetap bulat
plt.show()

# --- Bar Plot dengan Palet Warna Lebih Menarik ---
plt.figure(figsize=(10, 6))
sns.barplot(x=job_demand.index, y=job_demand, palette="coolwarm")

# Menambahkan judul dan label
plt.title('Top Job Titles by Demand', fontsize=16, fontweight='bold')
plt.xlabel('Job Title', fontsize=12)
plt.ylabel('Count of Job Titles', fontsize=12)

# Menambahkan nilai pada setiap bar
for index, value in enumerate(job_demand):
    plt.text(index, value + 1, f'{value}', ha='center', va='bottom', fontsize=12, color='black')

plt.xticks(rotation=45, ha='right')  # Agar label pada sumbu X terlihat lebih baik
plt.tight_layout()
plt.show()


# In[8]:


# Menghitung rata-rata gaji per Job Title dan Tingkat Pengalaman
salary_trend = position.groupby(['job_title', 'experience_level'])['avg_salary'].mean().unstack()
salary_trend = salary_trend.round(1)

# --- Heatmap Visualization ---
plt.figure(figsize=(12, 8))
sns.heatmap(salary_trend, annot=True, fmt=".1f", cmap="RdYlGn", linewidths=0.5, cbar_kws={'label': 'Average Salary ($)'})

# Menambahkan judul dan label
plt.title('Average Salary per Job Title and Experience Level', fontsize=16, fontweight='bold')
plt.xlabel('Experience Level', fontsize=12)
plt.ylabel('Job Title', fontsize=12)
# Menyesuaikan tata letak
plt.xticks(rotation=45, ha='right')
plt.tight_layout()

# Menampilkan plot
plt.show()

