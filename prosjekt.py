# -*- coding: utf-8 -*-
"""
@author: lars
"""

import pandas as pd #import for analyse av data
import matplotlib.pyplot as plt #import for plotting
clear2 = "\n" * 2 #legger 2 linjeskift til variabel clear5

#u_dag, k_slett, varighet, score = [], [], [], [] #globale data

def les_data(): #funksjon for å lese data. Del A
    global u_dag, k_slett, varighet, score #globale variabler
    df = pd.read_excel('support_uke_24.xlsx') #leser excel-filen
    u_dag = df.iloc[:, 0].tolist() #leser kolonne ukedag 
    k_slett = df.iloc[:, 1].tolist() #leser kolonne klokkeslett
    varighet = df.iloc[:, 2].tolist() #leser kolonne varighet
    score = df.iloc[:, 3].tolist() #leser kolonne tilfredshet
    print("Fil lest") #printer at filen er lest, for å være sikker på at filinnlesning gikk i orden
    
    #ett akutelt tillegg her kunne vært å spørre bruker om å taste inn filnavn, men det var ikke en del av oppgaven
    
def vis_antall_henvendelser_per_dag(): #funksjon for å plotte antall hendvendelser pr dag. Del B
    counts = pd.Series(u_dag).value_counts().sort_index() #teller antall pr ukedag
    counts.plot(kind='bar', title="Antall henvendelser per ukedag") #søylediagram(bar) med tittel
    plt.xlabel("Ukedag") #lager en etikett/label som gjør det enklere å lese diagrammet
    plt.ylabel("Antall henvendelser") #lager en etikett/label som gjør det enklere å lese diagrammet
    plt.show() #viser plottet

def finn_min_maks_varighet(): #funksjon for å finne og vise minimum og maksimum varighet på samtalene. Del C
    print(f"Minste samtaletid: {min(varighet)} ") #finner og printer  minste varighet
    print(f"Lengste samtaletid: {max(varighet)} ") #finner og printer lengste varighet

def gjennomsnitt_varighet(): #funksjon for å finne og vise gjennomsnittelig varighet på samtalene. Del D
#dette var en lur en, det tok en del tid å finne ut hvordan jeg skulle splitte og gjøre om til sekunder og tilbake igjen
    
    totale_sekunder = 0 #sørge for å starte på 0
    antall = 0 #sørge for å starte på 0
    
    # løkke for å beregne antall sekunder
    for tid in varighet:
        h, m, s = map(int, tid.split(':')) # dele opp i timer, minutter og sekunder
        sekunder = h * 3600 + m * 60 + s # konvertere til sekunder
        totale_sekunder += sekunder #akumulere sekunder
        antall = antall + 1 #øke antall med 1 for å holde orden på hvor mange "varigheter" det er
    
    snitt = totale_sekunder / antall #dele totale antall sekunder på antall "varigheter" for å få snittet
  
    # Formatere resultet til timer, minutter og sekunder igjen
    timer = snitt//3600 #finne timer
    rest = snitt%3600 #finne resten
    minutter = rest // 60 #finne minutter
    sekunder = rest % 60 #finne resten/sekunder
    
    formatert_tid = f"{timer:.0f}:{minutter:.0f}:{sekunder:.0f}" #formatere slik at det blir lett å lese
    print("Gjennomsnittelig varighet er:", formatert_tid) #skrive ut gjennomsnittelig varighet


def vis_henvendelser_per_tidsrom(): #funksjon for å vise antall henvendelser pr tidsrom. Del E
    tidsrom1 = 0 #variabel for antall mellom 8-10
    tidsrom2 = 0 #variabel for antall mellom 11-12
    tidsrom3 = 0 #variabel for antall mellom 13-14
    tidsrom4 = 0 #variabel for antall mellom 15-16
    klokke = 0 #variabel å ta imot integer/heltall fra stringvariabel ifm konvertering

    for tid in k_slett: #start løkke for å gå gjennom alle klokkeslett
        klokke = int(tid[:2]) #henter ut de 2 første tegnene fra tidspunkt, som er time, og konvertering til integer
        if 8 <= klokke <= 9: #sjekk om tidspunkt er mellom 8 og 9
            tidsrom1 += 1 #legge 1 til variabel tidsrom1 for å ta vare på antallet
        elif 10 <= klokke <= 11: #sjekk om tidspunkt er mellom 10 og 11
            tidsrom2 += 1 #legge 1 til variabel tidsrom2 for å ta vare på antallet
        elif 12 <= klokke <= 13: #sjekk om tidspunkt er mellom 12 og 13
            tidsrom3 += 1 #legge 1 til variabel tidsrom3 for å ta vare på antallet
        elif 14 <= klokke <= 15: #sjekk om tidspunkt er mellom 14 og 16
            tidsrom4 += 1 #legge 1 til variabel tidsrom4 for å ta vare på antallet

    verdier = tidsrom1, tidsrom2, tidsrom3, tidsrom4 #verdiene til plottet
    etiketter = ['08-10', '11-12', '13-14', '15-16'] #dette er etikettene for vaktene i plottet
    etiketter_med_antall = [f"{navn}: {verdi} stk" for navn, verdi in zip(etiketter, verdier)] # Slå sammen navn og antall i en etikett

# plotte Sektordiagram
    plt.pie(verdier, labels=etiketter_med_antall)
    plt.title("Antall henvendelser pr tidsrom/vakt")
    plt.axis('equal')
    plt.show()
    
def beregn_nps():
    positive = sum(1 for s in score if s >= 9) #positiv tilfredshet
    negative = sum(1 for s in score if 1 <= s <= 6) #negativ tilfredshet
    total = sum(1 for s in score if 1 <= s <= 10) #totalt antall
    nps = (positive / total * 100) - (negative / total * 100) #beregner tilfredsheten ved å trekke negative fra positive
    print(f"Net Promoter Score (NPS)er: {nps:.2f} prosent") #printer tilfredsheten
    

def vis_meny(): #definisjon av meny-funksjon
    while True: #løkke for å alltid vise menyen
        print (clear2) #printer variabel clear2 (linjeskiftene) for å gjøre konsollet mer lesbart 
        print("1. Antall henvendelser per ukedag (Del B)") #Del B
        print("2. Minste og lengste samtaletid (Del c)") #Del C
        print("3. Gjennomsnittlig samtaletid (Del d)") #Del D
        print("4. Henvendelser per tidsrom (Del e)") #Del E
        print("5. Beregn NPS (Del f)") #Del F
        print("6. Avslutt")
        valg = input("Velg et alternativ (1-6): ")
      
        if valg == '1':
            vis_antall_henvendelser_per_dag() #kaller funksjon vis_antall_henvendelser_per_dag hvis 1 er valgt
        elif valg == '2':
            finn_min_maks_varighet() #kaller funksjon finn_min_maks_varighet hvis 2 er valgt
        elif valg == '3':
            gjennomsnitt_varighet() #kaller funksjon gjennomsnitt_varighet hvis 3 er valgt
        elif valg == '4':
            vis_henvendelser_per_tidsrom() #kaller funksjon vis_henvendelser_per_tidsrom hvis 4 er valgt
        elif valg == '5':
            beregn_nps() #kaller funksjon beregn_nps hvis 5 er valgt
        elif valg == '6':
            print("Avslutter programmet.")
            break #avslutter programmet hvis 6 er valgt
        else:
            print("Ugyldig valg. Prøv igjen.") #gir melding hvis man taster ugyldig valg (ikke 1-6)

# dette er "hovedprogrammet" som først kaller les_data for å lese inn alle data fra filen, 
#for så å kalle vis_meny for å få bruker til å velge hva hen vil gjøre
les_data() #kaller funksjon les_data Del A
vis_meny() #kaller funksjon vis_meny
