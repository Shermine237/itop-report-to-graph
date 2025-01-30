import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, PieChart, LineChart, Reference

# Read and prepare data
def prepare_data(data):
    df = pd.read_csv(data)
    df['Date de début'] = pd.to_datetime(df['Date de début'])
    df['Date de fermeture'] = pd.to_datetime(df['Date de fermeture'])
    df['Date de résolution'] = pd.to_datetime(df['Date de résolution'])
    df['Mois'] = df['Date de début'].dt.strftime('%Y-%m')
    df['Délai de résolution'] = pd.to_numeric(df['Délai de résolution'], errors='coerce')
    return df

# Create workbook
wb = Workbook()

# 1. Tickets par état
def create_status_sheet(wb, df):
    ws = wb.create_sheet("État des tickets")
    status_count = df['Etat'].value_counts().reset_index()
    status_count.columns = ['État', 'Nombre de tickets']
    
    for r in dataframe_to_rows(status_count, index=False, header=True):
        ws.append(r)

    chart = PieChart()
    data = Reference(ws, min_col=2, min_row=1, max_row=len(status_count)+1)
    labels = Reference(ws, min_col=1, min_row=2, max_row=len(status_count)+1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.title = "Distribution des tickets par état"
    ws.add_chart(chart, "E2")

# 2. Tickets par catégorie
def create_category_sheet(wb, df):
    ws = wb.create_sheet("Catégories")
    cat_count = df['Sous catégorie de service->Nom'].value_counts().reset_index()
    cat_count.columns = ['Catégorie', 'Nombre de tickets']
    
    for r in dataframe_to_rows(cat_count, index=False, header=True):
        ws.append(r)

    chart = BarChart()
    data = Reference(ws, min_col=2, min_row=1, max_row=len(cat_count)+1)
    labels = Reference(ws, min_col=1, min_row=2, max_row=len(cat_count)+1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.title = "Tickets par catégorie de service"
    ws.add_chart(chart, "E2")

# 3. Analyse temporelle
def create_time_analysis_sheet(wb, df):
    ws = wb.create_sheet("Analyse temporelle")
    time_data = df.groupby('Mois')['Référence'].count().reset_index()
    time_data.columns = ['Mois', 'Nombre de tickets']
    
    for r in dataframe_to_rows(time_data, index=False, header=True):
        ws.append(r)

    chart = LineChart()
    data = Reference(ws, min_col=2, min_row=1, max_row=len(time_data)+1)
    labels = Reference(ws, min_col=1, min_row=2, max_row=len(time_data)+1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.title = "Évolution du nombre de tickets par mois"
    ws.add_chart(chart, "E2")

# 4. Analyse des délais de résolution
def create_resolution_time_sheet(wb, df):
    ws = wb.create_sheet("Délais de résolution")
    df['Délai (heures)'] = df['Délai de résolution'] / 3600
    resolution_stats = df.groupby('Sous catégorie de service->Nom')['Délai (heures)'].agg(['mean', 'median', 'min', 'max']).round(2).reset_index()
    
    for r in dataframe_to_rows(resolution_stats, index=False, header=True):
        ws.append(r)

    chart = BarChart()
    data = Reference(ws, min_col=2, min_row=1, max_row=len(resolution_stats)+1)
    labels = Reference(ws, min_col=1, min_row=2, max_row=len(resolution_stats)+1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.title = "Délai moyen de résolution par catégorie"
    ws.add_chart(chart, "E2")

# 5. Performance des agents
def create_agent_performance_sheet(wb, df):
    ws = wb.create_sheet("Performance agents")
    agent_stats = df.groupby('Agent->Nom complet')['Référence'].count().reset_index()
    agent_stats.columns = ['Agent', 'Nombre de tickets traités']
    
    for r in dataframe_to_rows(agent_stats, index=False, header=True):
        ws.append(r)

    chart = PieChart()
    data = Reference(ws, min_col=2, min_row=1, max_row=len(agent_stats)+1)
    labels = Reference(ws, min_col=1, min_row=2, max_row=len(agent_stats)+1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.title = "Distribution des tickets par agent"
    ws.add_chart(chart, "E2")

# Create summary sheet
def create_summary_sheet(wb, df):
    ws = wb.active
    ws.title = "Résumé"
    
    summary_data = [
        ["Nombre total de tickets", len(df)],
        ["Tickets résolus", len(df[df['Etat'] == 'Résolue'])],
        ["Délai moyen de résolution (heures)", df['Délai de résolution'].mean() / 3600],
        ["Catégorie la plus fréquente", df['Sous catégorie de service->Nom'].mode()[0]],
        ["Agent le plus actif", df['Agent->Nom complet'].mode()[0]]
    ]
    
    for row in summary_data:
        ws.append(row)

# Main execution
def create_analysis_workbook(data_file, output_file):
    df = prepare_data(data_file)
    
    create_summary_sheet(wb, df)
    create_status_sheet(wb, df)
    create_category_sheet(wb, df)
    create_time_analysis_sheet(wb, df)
    create_resolution_time_sheet(wb, df)
    create_agent_performance_sheet(wb, df)
    
    wb.save(output_file)

