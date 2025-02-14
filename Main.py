import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
from email.message import EmailMessage
import ssl
import smtplib

#Bewegungsdaten für Mai aus der Excel-Datei lesen
df = pd.read_excel(r"Bewegungsdaten\Kontoumsaetze_2023_05.xlsx",
                   dtype={"Art": "category"},
                   sheet_name="Fact_Data",
                   usecols=["Zeitstempel", "Kundennummer", "Art", "Betrag"])

#Berechnung neuer Spalten
df["Zeitstempel"] = pd.to_datetime(df['Zeitstempel'])
df["Datum"] = df['Zeitstempel'].dt.date
df["Monat"] = df['Zeitstempel'].dt.month
df["Ist_Einzahlung"] = df["Art"].map(
    lambda x: True if x in ["Bargeldeinzahlung", "Gehalt", "Überweisung (Ein)"] else False)


def ab_eingang(art):
    if art in ["Bargeldeinzahlung", "Gehalt", "Überweisung (Ein)"]:
        return 1
    else:
        return -1

def ist_bargeld_einzahlung(art):
    if art in ["Bargeldeinzahlung", "Überweisung (Ein)"]:
        return 1
    else:
        return 0

df["Betrag_Minus_Plus"] = df["Betrag"] * df["Art"].map(ab_eingang)
df["Bargeldeinzahlung"] = df["Betrag_Minus_Plus"].map(lambda x: 0 if float(x) <= 0 else float(x)) * df["Art"].map(ist_bargeld_einzahlung)

# 1-Monatiger kum. Galdeingang
df["KUM_GELDEINGANG"] = (df
.set_index('Zeitstempel')
.groupby('Kundennummer')['Bargeldeinzahlung']
.transform(lambda d: d.rolling('30D').sum())).reset_index()["Bargeldeinzahlung"]


#Limit-Überschreitung Validierung
df["Limit_Ueberschritten"] = df["KUM_GELDEINGANG"].map(lambda x: x if x >=10000  else 0)

#print(df[["Art","Betrag","Ist_Einzahlung","Betrag_Minus_Plus","Bargeldeinzahlung"]].tail(50))

#Kunden Infos
kunden_daten = pd.read_excel(
    r"Stammdaten\Kundendaten.xlsx",
    sheet_name="Kundendaten",
    usecols=["Kundennummer", "Vorname","Nachname", "Geschlecht", "Strasse","Hausnummer","PLZ","Wohnort","email"])

#Daten nach überschrittenen Geldeingängen filtern
ueberschrittene_geldeingaenge = df[df["Limit_Ueberschritten"]>=10000].drop_duplicates( subset=["Kundennummer"],
                                                                                       keep="last")

#Daten mit Kunden Infos joinen
ueberschrittene_geldeingaenge_infos = ueberschrittene_geldeingaenge.merge(kunden_daten,
                                                                          on="Kundennummer",
                                                                          how="left")

#BEstimmmte Spalten auswählen
ueberschrittene_geldeingaenge_emails = ueberschrittene_geldeingaenge_infos[["Kundennummer",
                                                                            "Vorname",
                                                                            "Nachname",
                                                                            "Geschlecht",
                                                                            "Strasse",
                                                                            "Hausnummer",
                                                                            "PLZ",
                                                                            "Wohnort",
                                                                            "email",
                                                                            "KUM_GELDEINGANG"]]


################################
###############################
##############################
# Email Schicken
def send_mail(empfaenger, geschlecht, name, kundennummer, betrag):
    absender = r"kimo.utube.69@gmail.com"
    kennwort = "qgsweepjcyioohyo"
    empfaenger = empfaenger

    grenze = 10000
    betreff = f"MusterBank (Nachweispflicht bei Bareinzahlungen über {grenze:,.2f} €) Kundennummer: {kundennummer}"
    anrede = "geehrter Herr" if geschlecht.upper() == "M" else "geehrte Frau"
    inhalt = f""" Sehr {anrede} {name},

    Seit dem 8. August 2021 gelten neue Regeln der Finanzaufsicht BaFin. 
    Bei Bargeld-Einzahlungen über 10.000 Euro müssen Banken und Sparkassen von Kunden einen sogenannten Herkunftsnachweis verlangen.

    Sie haben in der letzten Zeit {betrag:,.2f} € auf Ihrem Konto eingezahlt. Aus diesem Grund sind Sie uns verpflichtet, einen Nachweis über die Herkunft des Gelds innerhalb von 2 Wochen einzureichen.

    Im Online Banking:
    1) Melden Sie sich mit Ihren Zugangsdaten an, um Ihren Nachweis bei Bareinzahlungen einreichen zu können.
    2) Füllen Sie die Felder entsprechend aus.
    3) Wählen Sie die getätigte Transaktion, z. B. "Bareinzahlung Euro", den Tag sowie den Betrag der Einzahlung aus.
    4) Wählen Sie die Art des Herkunftsnachweises aus, z. B. "Barauszahlungsquittung" und laden Sie den entsprechenden Nachweis hoch.
    5) Mit Klick auf "Weiter" sehen Sie eine Zusammenfassung Ihrer eingegebenen Daten.
    6) Klicken Sie auf "Senden", um das Einreichen Ihres Nachweises freizugeben.


    Aussagekräftige Belege nach Auskunft der BaFin:
    ◉ Aktueller Kontoauszug des Kundenkontos bei einer anderen Bank.
    ◉ Barauszahlungsquittungen einer anderen Bank oder Sparkasse.
    ◉ Sparbuch, aus dem die Barauszahlung hervorgeht.
    ◉ Verkaufs- und Rechnungsbelege (z. B. Belege zu einem Auto-, oder Warenverkauf, Verkauf von Dienstleistungen).
    ◉ Letztwillige vom Nachlassgericht eröffnete Verfügungen, Erbschein oder ähnliche Erbnachweise.
    ◉ Schenkungsverträge oder Schenkungsanzeigen.

    Hinweis:
    Die Auflistung der Belege ist nicht abschließend. In Einzelfällen ist zu prüfen, ob ein vorgelegter Nachweis ausreichend ist und die Herkunft plausibel dargelegt werden kann.

    Mit freundlichen Grüßen

    Ihre MusterBank
    """

    em = EmailMessage()

    em["From"] = absender
    em["To"] = empfaenger
    em["subject"] = betreff
    em.set_content(inhalt)

    context = ssl.create_default_context()

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as smtp:
        smtp.login(absender, kennwort)
        smtp.sendmail(absender, empfaenger, em.as_string())

#Email an diese Kunden schicken
for index,i in enumerate(ueberschrittene_geldeingaenge_emails.to_numpy()):
    email = i[8]
    geschlecht = i[3]
    nachname = i[2]
    kundennummer = i[0]
    betrag = i[9]
    send_mail(email,geschlecht,nachname,kundennummer, betrag) if  index %2 ==0 else 0

#heutiges Datum berechnen
heute = pd.to_datetime('today').date()

#Ordner für die Visualisierungen erstellen
path = rf"Ergebnisdaten\{heute}"
isExist = os.path.exists(path)
if not isExist:
   os.makedirs(path)


#Ergebnis in Excel speichern
ueberschrittene_geldeingaenge_emails.to_excel(
    path+rf"\Kontos mit ueberschrittenen Geldeingängen am {heute}.xlsx",
    index=False)

#Kontoanalyse nach überschrittenen Konten filtern
umsaetze_ueberschritten = df[df["Kundennummer"].isin(ueberschrittene_geldeingaenge["Kundennummer"])]

#Kontoanalyse nur Zeitstempel, Kundennummer und Betrag
umsaetze_ueberschritten = umsaetze_ueberschritten[["Zeitstempel","Kundennummer","Betrag_Minus_Plus"]]

#Kontostände am Anfang des Monats laden und eine Dummy-Spalte für Zeitstempel erstellen
Kontostaende_ueberschritten = pd.read_excel(
    r"Bewegungsdaten\Kontostaende_2023_05.xlsx",
    sheet_name="Kontostand",
    usecols=["Kundennummer", "Kontostand"]).rename({"Kontostand" : "Betrag_Minus_Plus"},axis=1)

Kontostaende_ueberschritten["Zeitstempel"] = pd.to_datetime("05.01.2023  00:00:01")

#Kontostände nach überschrittenen Konten filtern
Kontostaende_ueberschritten= Kontostaende_ueberschritten[Kontostaende_ueberschritten["Kundennummer"].isin(ueberschrittene_geldeingaenge["Kundennummer"])]


#Spalten neu anordnen
Kontostaende_ueberschritten = Kontostaende_ueberschritten.iloc[:,[2,0,1]]
umsaetze_ueberschritten = umsaetze_ueberschritten.iloc[:,[0,1,2]]

#Kontostände und Kontoumsätze afür überschrittene aneinander fügen
datenanaylse_ueberschritten = pd.concat([Kontostaende_ueberschritten,umsaetze_ueberschritten],
                      ignore_index=True,
                      axis=0)


#Kum. Summe für die Datenanalyse
datenanaylse_ueberschritten["Kontostand"] = datenanaylse_ueberschritten.groupby(["Kundennummer"])["Betrag_Minus_Plus"].cumsum()
datenanaylse_ueberschritten["Datum"] = datenanaylse_ueberschritten['Zeitstempel'].dt.date

#Ergebnis speichern
datenanaylse_ueberschritten.to_excel(
    path+rf"\Datenanalyse für die überschrittenen Konten (Uhrzeitsweise) am {heute}.xlsx"
    ,index=False)


datenanaylse_ueberschritten_plot = datenanaylse_ueberschritten.drop_duplicates(subset=["Datum","Kundennummer"],
                                                                               keep="last",
                                                                               ignore_index=True)
datenanaylse_ueberschritten_plot = datenanaylse_ueberschritten_plot[["Kundennummer","Datum","Kontostand"]]

#Ergebnis speichern
datenanaylse_ueberschritten_plot.to_excel(
    path+rf"\Datenanalyse für die überschrittenen Konten (Tagesweise) am {heute}.xlsx",
    index=False)

#Ordner für die Visualisierungen erstellen
path = rf"Ergebnisvisualisierungen\{heute}"
isExist = os.path.exists(path)
if not isExist:
   os.makedirs(path)

#Daten für 5 zufällig ausgewählte Kunden visualisieren
anzahl_u_konten = len(datenanaylse_ueberschritten_plot["Kundennummer"].drop_duplicates())
n= 5
kunden_ueberschritten_5 = datenanaylse_ueberschritten_plot["Kundennummer"].drop_duplicates().sample(n=n).to_numpy()
data =  datenanaylse_ueberschritten_plot[datenanaylse_ueberschritten_plot["Kundennummer"].isin(kunden_ueberschritten_5)]
plt.figure(figsize=(14,7))
sns.lineplot(x= data["Datum"],
             y=data["Kontostand"],
             hue=data["Kundennummer"])

plt.legend()
plt.title(f"Kontoanalyse am Tagesende\n (Sample = {n} von {anzahl_u_konten} überschrittenen Konten)")
plt.savefig(path+rf"\Kontoanalyse (Sample = {n}) am {heute}")
plt.show()