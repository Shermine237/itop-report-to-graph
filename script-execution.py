
# Import du script d'analyse
from analyse_tickets import create_analysis_workbook

# Exécution
if __name__ == "__main__":
    
    # Créer le fichier Excel avec les analyses
    # create_analysis_workbook(csv_file, 'analyse_tickets.xlsx')
    create_analysis_workbook('Export de Ticket.csv', 'analyse_tickets.xlsx')
    print("Analyse terminée ! Le fichier 'analyse_tickets.xlsx' a été créé.")
