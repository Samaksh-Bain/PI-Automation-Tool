from flask import Flask, render_template, request
import pandas as pd
import requests
from flask import Flask, render_template, request, send_file
import pandas as pd
from collections import defaultdict
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
import numpy as np
import io

app = Flask(__name__)

# Load data once when the app starts
sheet_name = 'Base year_2021'
usecols = 'B:AK'
df = pd.read_excel('20230411_A1OI_Summary Update_v4.xlsx', sheet_name=sheet_name, usecols=usecols)
priority_companies_df = pd.read_excel('A1OI Priority client mapping.xlsx')
priority_companies = priority_companies_df['Informal client name'].tolist()
df['Priority_status'] = np.where(df['Company'].isin(priority_companies_df['Informal client name']), 'Priority', 'Non-Priority')
df['Geography'] = df['Geography'].replace('N. America', 'AMER')
df = df.fillna(0)

def get_contextual_values(df, geography=None, industry=None):
    if geography:
        filtered_df = df[df['Geography'] == geography]
    else:
        filtered_df = df
    
    if industry:
        filtered_df = filtered_df[filtered_df['Bain Industry'] == industry]
    
    geographies = sorted(df['Geography'].unique().tolist())
    countries = sorted(filtered_df['Country'].unique().tolist()) if geography else sorted(df['Country'].unique().tolist())
    industries = sorted(df['Bain Industry'].unique().tolist())
    industries_primary = sorted(filtered_df['Primary Industry'].unique().tolist())
    relationships = sorted(df['Bain Relationship'].unique().tolist())
    priority_column = sorted(df['Priority_status'].unique().tolist())

    return industries, relationships, countries, geographies, industries_primary, priority_column

def filter_data(df, industry, relationship, revenue, geography, country, industry_primary, combined_opp_score, priority_status):
    df['Country'] = df['Country'].str.strip()

    if isinstance(country, str):
        country = [c.strip() for c in country.split(',')]
    else:
        country = [c.strip() for c in country]

    if industry:
        df = df[df['Bain Industry'] == industry]
    if industry_primary:
        df = df[df['Primary Industry'] == industry_primary]
    if relationship:
        if relationship == 'Current':
            df = df[df['Bain Relationship'].str.startswith('Current')]
        elif relationship == 'Recent':
            df = df[df['Bain Relationship'].str.startswith('Recent')]
        elif relationship == 'Greyspace':
            df = df[df['Bain Relationship'].str.startswith('Greyspace')]
    if revenue == 'less_than_500000':
        df = df[df['Revenue (M)'] < 5000]
    elif revenue == 'greater_than_500000':
        df = df[df['Revenue (M)'] > 5000]
    if priority_status:
        df = df[df['Priority_status'] == priority_status]
    if geography:
        df = df[df['Geography'] == geography]
    if country:
        df = df[df['Country'].isin(country)]
    if combined_opp_score:
        if combined_opp_score == 'less_than_5':
            df = df[df['Combined Opp. Score'] < 5]
        elif combined_opp_score == 'between_5_and_10':
            df = df[(df['Combined Opp. Score'] >= 5) & (df['Combined Opp. Score'] <= 10)]
        elif combined_opp_score == 'greater_than_10':
            df = df[df['Combined Opp. Score'] > 10]
    
    return df

@app.route('/')
def home():
    industry = request.args.get('industry')
    relationship = request.args.get('relationship')
    revenue = request.args.get('revenue')
    country = request.args.getlist('country[]')
    geography = request.args.get('geography')
    industry_primary = request.args.get('industry_primary')
    combined_opp_score = request.args.get('combined_opp_score')
    priority_status = request.args.get('priority_status')

    filters_applied = any([industry, relationship, revenue, country, geography, industry_primary, combined_opp_score, priority_status])

    df_filtered = filter_data(df, industry, relationship, revenue, geography, country, industry_primary, combined_opp_score, priority_status)
    filtered_data = df_filtered.to_dict(orient='records')
    min_revenue = df_filtered['Revenue (M)'].min()

    industries, relationships, countries, geographies, industries_primary, priority_column = get_contextual_values(df, geography, industry)

    return render_template('index.html',
                           industries=industries,
                           industries_primary=industries_primary,
                           relationships=relationships,
                           geographies=geographies,
                           countries=countries,
                           revenue_ranges=['less_than_500000', 'greater_than_500000'],
                           filtered_data=filtered_data,
                           filters_applied=filters_applied,
                           industry=industry,
                           industry_primary=industry_primary,
                           revenue=revenue,
                           relationship=relationship,
                           geography=geography,
                           country=country,
                           min_revenue=min_revenue,
                           priority_companies=priority_companies,
                           priority_column=priority_column,
                           priority_status=priority_status)

@app.route('/generate_ppt', methods=['POST'])
def generate_ppt():
    post_url = 'http://localhost:5001/generate_ppt'
    response = requests.post(post_url, data=request.form)

    return send_file(
        io.BytesIO(response.content),  # Content as bytes
        download_name="output.pptx",   # The name of the file when downloaded
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",  # MIME type for PPTX
        as_attachment=True  # Force download
    )
if __name__ == '__main__':
    app.run(debug=True, port=5000)
