""" import click
from modules.modules import create_monthly_schedule
@click.command()
@click.option('--dateiname', '-d', help='Name der Excel-Datei (ohne .xlsx)')
@click.option('--zuschlag', '-z', is_flag=True, help='Aktiviert den Feiertagszuschlag')
@click.option('--gehalt', '-g', type=float, default=12.50, help='Stundenlohn (Standard: 12.50)')
@click.option('--bundesland', '-bl', default='SN', help='Auswahl des Bundeslandes (für die Feiertage), Standard ist Sachsen (1)')

def main(dateiname, bundesland, zuschlag, gehalt):
    
    erstellte_datei = create_monthly_schedule(dateiname, bundesland, zuschlag, gehalt)
    print(f"Excel-Datei wurde erstellt: {erstellte_datei}")

if __name__ == "__main__":
    main() """
    
from modules.gui import MonthlyScheduleApp

def main():
    app = MonthlyScheduleApp()
    app.run()

if __name__ == "__main__":
    main()