# Web εφαρμογή αναζήτησης ΓΕΜΗ

Η εφαρμογή έχει πεδία για:

- ειδικότητα / κλάδο,
- περιοχή,
- ΚΑΔ (προαιρετικά),
- Δήμο ή Νομό,
- ενεργές επιχειρήσεις,
- μέγιστο αριθμό αποτελεσμάτων.

Τα αποτελέσματα εμφανίζονται σε πίνακα και κατεβαίνουν σε Excel.

## Τοπική εκτέλεση σε Windows

1. Εγκατέστησε Python.
2. Αντέγραψε το `.streamlit/secrets.toml.example` ως `.streamlit/secrets.toml`.
3. Βάλε μέσα το νέο API key.
4. Κάνε διπλό κλικ στο `run_windows.bat`.
5. Η εφαρμογή ανοίγει συνήθως στη διεύθυνση `http://localhost:8501`.

Εναλλακτικά, σε PowerShell:

```powershell
py -m pip install -r requirements.txt
py -m streamlit run app.py
```

## Ανέβασμα στο Streamlit Community Cloud

1. Δημιούργησε νέο repository στο GitHub.
2. Ανέβασε τα αρχεία του φακέλου, αλλά όχι πραγματικό `secrets.toml`.
3. Στο Streamlit Community Cloud πάτησε **Create app** και επίλεξε το repository και το `app.py`.
4. Στο **Advanced settings → Secrets** βάλε:

```toml
GEMI_API_KEY = "ΤΟ_ΝΕΟ_API_KEY"
```

5. Πάτησε **Deploy**.

## Ασφάλεια

- Μην γράψεις το API key μέσα στο `app.py`.
- Μην ανεβάσεις το πραγματικό `.streamlit/secrets.toml` στο GitHub.
- Το API του ΓΕΜΗ έχει όριο 8 αιτημάτων ανά λεπτό. Η εφαρμογή περιμένει αυτόματα ανάμεσα στις κλήσεις.
- Χρησιμοποίησε νέο API key, επειδή το προηγούμενο έχει κοινοποιηθεί στη συνομιλία.

## Πηγή δεδομένων

OpenData ΓΕΜΗ — άδεια ODC-BY-1.0.
