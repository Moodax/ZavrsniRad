# Pretvarač Excel u CSV s DCAT Metapodacima

Python paket za pretvaranje Excel datoteka u CSV s inteligentnim prepoznavanjem tablica i generiranjem DCAT-AP metapodataka.

## Značajke

- Inteligentno prepoznavanje tablica u Excel datotekama
- Podrška za spojene ćelije i složene rasporede
- Obrada u memoriji za veću učinkovitost
- Generiranje DCAT-AP usklađenih metapodataka
- CLI i GUI sučelja

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

### Grafičko korisničko sučelje (GUI)

Pokrenite GUI za interaktivno iskustvo:

```bash
excel_to_csv_dcat_gui
```

### Python API

Paket možete koristiti i programatski u svojim Python skriptama:

```python
from excel_to_csv_dcat.core import extract_tables_from_excel

with open("input.xlsx", "rb") as f:
    excel_bytes = f.read()

tables = extract_tables_from_excel(excel_bytes)

for name, csv_buffer in tables:
    with open(f"{name}.csv", "wb") as out_file:
        out_file.write(csv_buffer.getvalue())
```

### Postavljanje okruženja

1. Klonirajte repozitorij:

   ```bash
   git clone https://github.com/Moodax/ZavrsniRad
   cd excel-to-csv-dcat
   ```

2. Instalirajte dependencies:

   ```bash
   pip install -r requirements.txt
   ```
