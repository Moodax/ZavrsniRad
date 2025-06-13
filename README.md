# Pretvarač Excel u CSV s DCAT Metapodacima

Python paket za pretvaranje Excel datoteka u CSV s inteligentnim prepoznavanjem tablica i generiranjem DCAT-AP metapodataka.

## Značajke

- Inteligentno prepoznavanje tablica u Excel datotekama
- Podrška za spojene ćelije i složene rasporede
- Obrada u memoriji za veću učinkovitost
- Generiranje DCAT-AP usklađenih metapodataka
- CLI i GUI sučelja
- **AI-powered funkcionalnosti**:
  - Automatsko generiranje zaglavlja za tablice bez zaglavlja
  - AI-podržana validacija tipova podataka
  - Integracija s OpenAI GPT i Google Gemini modelima

## AI Funkcionalnosti

Alat sada podržava napredne AI funkcionalnosti za poboljšanje kvalitete podataka:

### Podržani AI Pružatelji
- **OpenAI GPT** (GPT-3.5-turbo)
- **Google Gemini** (Gemini-1.5-flash)

### AI Značajke
1. **Generiranje zaglavlja**: AI analizira sadržaj stupaca i predlaže smislena imena zaglavlja
2. **Validacija tipova podataka**: AI pomaže u prepoznavanju i validaciji tipova podataka za CSVW metapodatke

### Postavljanje AI Funkcionalnosti

Da biste koristili AI funkcionalnosti, trebat ćete API ključ za odabranog pružatelja (OpenAI ili Google Gemini). Ovaj ključ se prosljeđuje alatu putem `--llm-api-key` opcije u CLI-u ili odgovarajućeg polja u GUI-u.

Za lakše upravljanje vašim API ključevima, možete ih pohraniti u `.env` datoteku u korijenu projekta (kreirajte je ako ne postoji, npr. kopiranjem `.env.example` ako postoji). Alat **neće automatski** čitati ove varijable; i dalje trebate eksplicitno proslijediti ključ putem argumenta. Primjer sadržaja `.env` datoteke:

```bash
# Primjer za pohranu OpenAI API ključa
OPENAI_API_KEY="vaš_openai_api_ključ"

# Primjer za pohranu Google Gemini API ključa
GOOGLE_API_KEY="vaš_google_api_ključ"
```

Kada pokrećete alat, koristite vrijednost spremljenog ključa za `--llm-api-key` opciju.

## Instalacija

Slijedite ove korake za instalaciju i pokretanje paketa na Windows operativnom sustavu:

1. **Klonirajte repozitorij**:

   - Otvorite PowerShell i upišite sljedeće naredbe:

     ```powershell
     git clone https://github.com/Moodax/ZavrsniRad
     cd excel-to-csv-dcat
     ```

2. **Kreirajte virtualno okruženje**:

   - Upišite sljedeću naredbu za kreiranje virtualnog okruženja:

     ```powershell
     python -m venv .venv
     ```

3. **Aktivirajte virtualno okruženje**:

   - Upišite sljedeću naredbu za aktivaciju virtualnog okruženja:

     ```powershell
     .venv\Scripts\activate
     ```

4. **Instalirajte dependencies**:

   - U virtualnom okruženju instalirajte potrebne dependencies:

     ```powershell
     pip install -r requirements.txt
     pip install -e .
     ```

5. **Pokrenite alat**:

   - **Za korištenje CLI sučelja**:

     ```powershell
     excel_to_csv_dcat input.xlsx -o output_dir -m turtle
     ```

   - **Za pokretanje GUI sučelja**:

     ```powershell
     excel_to_csv_dcat_gui
     ```

### Sučelje naredbenog retka (CLI)

Pretvorite Excel datoteku u CSV i generirajte metapodatke pomoću CLI-a:

```bash
excel_to_csv_dcat input.xlsx -o output_dir -m turtle
```

#### Opcije

- `-o`, `--output-dir`: Izlazni direktorij za CSV datoteke (zadano: `output`)
- `-m`, `--metadata-format`: Format metapodataka (`turtle` ili `json-ld`)
- `-b`, `--base-uri`: Osnovni URI za identifikatore skupova podataka
- `-p`, `--publisher-uri`: URI izdavača
- `-n`, `--publisher-name`: Naziv izdavača
- `-l`, `--license`: URI licence
- `--existing-dataset-uri`: Opcionalni URI postojećeg DCAT Dataseta na koji će se povezati generirane distribucije.
- `--original-source-description`: Opcionalni tekst koji opisuje originalni izvor dataseta.

#### AI Opcije

- `--enable-ai`: Omogući AI funkcionalnosti
- `--llm-provider`: Odaberite AI pružatelja (`openai` ili `gemini`)
- `--llm-api-key`: API ključ za odabrani AI pružatelj
- `--skip-header-ai`: Preskočite AI generiranje zaglavlja
- `--skip-datatype-ai`: Preskočite AI validaciju tipova podataka

#### Primjer korištenja s AI funkcionalnostima

```bash
# Korištenje OpenAI za AI funkcionalnosti
excel_to_csv_dcat input.xlsx -o output_dir --enable-ai --llm-provider openai --llm-api-key your_api_key

# Korištenje Gemini samo za zaglavlja
excel_to_csv_dcat input.xlsx -o output_dir --enable-ai --llm-provider gemini --skip-datatype-ai
```

### Grafičko korisničko sučelje (GUI)

Pokrenite GUI za interaktivno iskustvo:

```bash
excel_to_csv_dcat_gui
```

### Python API

Paket možete koristiti i programatski u svojim Python skriptama:

```python
from excel_to_csv_dcat.core import extract_tables_from_excel
import os 

with open("input.xlsx", "rb") as f:
    excel_bytes = f.read()

extracted_table_details = extract_tables_from_excel(excel_bytes)

# Primjer spremanja CSV datoteka u trenutni direktorij
# Možete specificirati i izlazni direktorij, npr. os.makedirs("output_api", exist_ok=True)
for table_detail in extracted_table_details:
    table_name = table_detail["name"]
    csv_buffer = table_detail["buffer"]  # BytesIO objekt

    # Odredišna putanja za CSV datoteku
    output_csv_path = f"{table_name}.csv" 
    # Ako želite spremiti u poddirektorij:
    # output_csv_path = os.path.join("output_api", f"{table_name}.csv")

    with open(output_csv_path, "wb") as out_file:
        out_file.write(csv_buffer.getvalue())

    print(f"Spremljena tablica: {output_csv_path}")
```
