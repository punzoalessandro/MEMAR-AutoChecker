# MEMAR AutoChecker

# Controllo Automatico dei Totali MEMAR

Questo progetto automatizza il controllo giornaliero dei totali MEMAR, semplificando il processo di verifica tra i dati generati internamente da **CSE** e quelli ricevuti da **Caricese**.

## Descrizione

Il software verifica automaticamente la corrispondenza tra i file MEMAR, eliminando il processo manuale tradizionalmente soggetto a errori. Grazie all'automazione, il controllo risulta più rapido, preciso ed efficiente.

### Benefici

- **Riduzione degli errori**: L'automazione previene gli errori umani nel confronto dei file.
- **Risparmio di tempo**: Il processo è significativamente più veloce rispetto alla verifica manuale.
- **Monitoraggio continuo**: Grazie a Watchdog, il software può monitorare automaticamente l'arrivo di nuovi file.

## Funzionalità

- Caricamento e confronto dei file MEMAR (formato Excel).
- Notifica automatica di eventuali discrepanze.
- Monitoraggio in tempo reale di una directory per rilevare nuovi file da controllare.
- Generazione di report dettagliati sull'esito del confronto.

## Requisiti

- **Python** 3.8 o superiore
- Librerie Python richieste:
  - `openpyxl` per la gestione dei file Excel.
  - `watchdog` per il monitoraggio delle directory.

## Installazione

1. Clona il repository:
   ```bash
   git clone https://github.com/tuo-username/totali-memar.git
   cd totali-memar
2. Installa le dipendenze:
   ```bash
   pip install -r requirements.txt

## Utilizzo
1. Avvia lo script principale:
   ```bash
   python main.py

## Contribuzione
Se hai idee per migliorare il progetto o riscontri problemi, sentiti libero di aprire una Issue o inviare una Pull Request. 


   
