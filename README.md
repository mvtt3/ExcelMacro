# 🧠 Matteo Pepe - Excel Macro Tools

<p>
  <a href="https://learn.microsoft.com/it-it/office/vba/api/overview/excel"><img src="https://img.shields.io/badge/code-VBA-blue" alt="VBA"></a>
  <a href="#"><img src="https://img.shields.io/badge/app-Microsoft_Excel-green" alt="Excel"></a>
  <a href="#"><img src="https://img.shields.io/badge/function-Data%20Cleaning-important" alt="Data Cleaning"></a>
  <a href="#"><img src="https://img.shields.io/badge/UI-Form%20based-lightgrey" alt="Userform"></a>
  <a href="#"><img src="https://img.shields.io/badge/version-1.0-orange" alt="Version"></a>
</p>

---

## ✨ Benvenuto!

Questa raccolta di **macro Excel in VBA** è stata sviluppata per **automatizzare il confronto, la pulizia e l’estrazione dei dati** tra più fogli Excel, in particolare per la verifica delle **Partite IVA** tra elenchi.

Se sei arrivato qui dal mio sito o portfolio, benvenuto nel cuore tecnico del progetto! 💻

---

## 📌 Funzionalità principali

- ✅ Confronto tra due elenchi di Partite IVA (anche parziali)  
- ✅ Evidenziazione automatica delle corrispondenze  
- ✅ Riconoscimento di duplicati nello stesso foglio  
- ✅ Esportazione in nuovo file con righe filtrate  
- ✅ Interfaccia semplice tramite UserForm  
- ✅ Logica smart che ignora caratteri non numerici nelle P.IVA  

---

## 🛠️ Tecnologie Utilizzate

| Linguaggio / Strumento | Utilizzo                                     |
|------------------------|----------------------------------------------|
| **VBA (Visual Basic)** | Logica delle macro, interazione file         |
| **Microsoft Excel**    | Interfaccia utente, struttura dei dati       |
| **UserForm**           | Selezione file e gestione flusso             |

➡️ Tutto è integrato in **un singolo file Excel abilitato alle macro** (`.xlsm`), pronto all’uso senza necessità di installazioni aggiuntive.

---

## 🧩 Architettura del Progetto

### 📦 File Principale

📄 `MatchPartiteIVA.xlsm`  
├─ Modulo: `ConfrontaPartiteIVA()` → Evidenzia match tra due file  
├─ Modulo: `EvidenziaDuplicati()` → Trova duplicati nello stesso elenco  
├─ Modulo: `EstraiRigheConMatch()` → Estrae solo righe con match  
└─ UserForm: Selezione file e avvio operazioni

---

## 🔄 Flusso di Interazione

1. L’utente apre il file Excel `.xlsm`  
2. Tramite **UserForm**, carica i due file da confrontare  
3. Avvia la macro `ConfrontaPartiteIVA()`  
4. I risultati vengono **colorati automaticamente** nel foglio  
5. Può poi:
   - Evidenziare duplicati nello stesso file
   - Esportare solo le righe corrispondenti in un nuovo file

---

## 🧪 Come Usarlo

### ▶️ Esecuzione

1. Scarica il file: `MatchPartiteIVA.xlsm`  
2. Abilita le macro all’apertura  
3. Segui la guida a video del **UserForm**  
4. Carica i due file Excel da confrontare  
5. Avvia le azioni desiderate

> 💡 Il file funziona **offline al 100%** e non richiede connessione a internet!

---

## 📷 Screenshot

### Interfaccia principale dello sviluppo in VBA

![UserForm Screenshot](https://i.imgur.com/5BrXwZ0.png)

### Esempio di confronto con celle evidenziate

![Risultato Match](https://i.imgur.com/BsboaWH.png)

---

## 🚀 Possibili Sviluppi Futuri

- Supporto per altri codici fiscali o ID cliente  
- Interfaccia utente migliorata con progress bar  
- Log automatico dei risultati in un file separato  
- Estensione per Excel Online con Office Script  

---

## 📬 Contatti

- 📧 Email: [matteopepe1701@gmail.com](mailto:matteopepe1701@gmail.com)  
- 📞 Telefono: +39 340 98 53 600  
- 💼 LinkedIn: [linkedin.com/in/matteo-pepe-a56477249](https://www.linkedin.com/in/matteo-pepe-a56477249/)

---

## 💬 Feedback

Hai suggerimenti o funzionalità da proporre?  
Apri una [Issue](https://github.com/tuo-username/excel-vba-tools/issues) o scrivimi direttamente!

---

> Progetto sviluppato con precisione e automazione da **Matteo Pepe** 💙
