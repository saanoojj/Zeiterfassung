import click
from modules.modules import create_monthly_schedule

@click.command()
@click.option('--dateiname', '-d', help='Name der Excel-Datei (ohne .xlsx)')
@click.option('--zuschlag', '-z', is_flag=True, help='Aktiviert den Feiertagszuschlag')
@click.option('--gehalt', '-g', type=float, default=12.50, help='Stundenlohn (Standard: 12.50)')
def main(dateiname, zuschlag, gehalt):
    erstellte_datei = create_monthly_schedule(dateiname, zuschlag, gehalt)
    print(f"Excel-Datei wurde erstellt: {erstellte_datei}")

if __name__ == "__main__":
    main()