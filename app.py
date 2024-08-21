import streamlit as st
import pandas as pd
import os
import matplotlib.pyplot as plt

def summarize_excel(file):
    xls = pd.ExcelFile(file)
    summary_by_feature = {}
    
    for sheet_name in xls.sheet_names:
        if sheet_name == "MoM_volume":
            continue
        
        df = pd.read_excel(xls, sheet_name=sheet_name)
        feature_name = df.columns[0]
        months = df.columns[1:]
        categories = df.iloc[:, 0]
        
        if len(months) < 2:
            continue
        
        latest_month = months[-1]
        previous_month = months[-2]
        
        summary_points = []
        display_df = False
        
        for i, category in enumerate(categories):
            try:
                latest_percentage = df.iloc[i, -1]
                previous_percentage = df.iloc[i, -2]
            except IndexError:
                continue
            
            difference = latest_percentage - previous_percentage
            if abs(difference) > 5:
                display_df = True
                trend = 'increased' if difference > 0 else 'decreased'
                deviation_number = f"{abs(difference):.2f}"
                
                summary_point = (
                    f"- For category **{category}**, number of accounts "
                    f"<u>**{trend}**</u> by {deviation_number}% in {latest_month} "
                    f"as compared to the previous month."
                )
                
                summary_points.append(summary_point)
        
        if summary_points:
            summary_by_feature[feature_name] = {
                "summary_points": summary_points,
                "display_df": display_df,
                "df": df[[df.columns[0]] + list(df.columns[-5:])] if len(categories) <= 6 else None
            }
    
    return summary_by_feature

def plot_mom_volume(file):
    df = pd.read_excel(file, sheet_name="MoM_volume")
    plt.figure(figsize=(10, 4))
    plt.plot(df['month_year'], df['account_count'], marker='o', linestyle='-', color='#007acc')
    #plt.title('Month on Month Volume Trend', fontsize=18, fontweight='bold')
    plt.xlabel('Month-Year', fontsize=14)
    plt.ylabel('Account Count', fontsize=14)
    plt.xticks(rotation=45)
    plt.grid(True, linestyle='--', alpha=0.7)
    st.pyplot(plt)

# Streamlit App
st.sidebar.title("Upload Excel Files")
st.sidebar.write("**Note:** You can upload multiple Excel files at a time.")
uploaded_files = st.sidebar.file_uploader("Choose Excel files", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    for uploaded_file in uploaded_files:
        filename = os.path.splitext(uploaded_file.name)[0]
        
        st.markdown(f"""
            <style>
            .banner {{
                background-color: #007acc;
                color: #ffffff;
                padding: 15px;
                text-align: center;
                border-radius: 8px;
                margin-bottom: 20px;
            }}
            .filename {{
                font-size: 28px;
                font-weight: bold;
            }}
            </style>
            <div class="banner">
                <div class="filename">{filename}</div>
            </div>
        """, unsafe_allow_html=True)
        
        st.markdown("<h2 style='font-size:24px;'>Month on Month Volume Trend</h2>", unsafe_allow_html=True)
        plot_mom_volume(uploaded_file)
        
        st.markdown("<h2 style='font-size:24px;'>Summary</h2>", unsafe_allow_html=True)
        summary_data = summarize_excel(uploaded_file)
        for feature_name, data in summary_data.items():
            st.markdown(f"**{feature_name}:**", unsafe_allow_html=True)
            st.markdown("\n".join(data["summary_points"]), unsafe_allow_html=True)
            
            if data["display_df"] and data["df"] is not None:
                #st.markdown(f"**{feature_name} DataFrame:**")
                st.dataframe(data["df"].style.set_properties(**{
                    'text-align': 'left',
                    'background-color': '#f4f4f4',
                    'color': '#333',
                    'border-color': 'black',
                    'border-style': 'solid',
                    'border-width': '1px',
                }))
        
        st.markdown("<hr>", unsafe_allow_html=True)
    
    st.markdown("""
        <style>
        .sidebar .sidebar-content {
            background-color: #f4f4f4;
            padding: 20px;
        }
        .main {
            font-family: 'Arial', sans-serif;
            line-height: 1.6;
            color: #333;
            padding: 20px;
        }
        h2 {
            color: #007acc;
        }
        .dataframe {
            margin-top: 20px;
        }
        hr {
            border: none;
            border-top: 2px solid #007acc;
            margin: 30px 0;
        }
        </style>
    """, unsafe_allow_html=True)
else:
    st.sidebar.write("Please upload Excel files to see the summaries.")








