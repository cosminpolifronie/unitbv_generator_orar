# unitbv_generator_orar
Scriptul extrage și formatează orarul Universității Transilvania pentru un număr nelimitat de grupe folosind orarul facultății în format general. [Rezultatul](https://i.imgur.com/ZJ3A6WH.png) este un fișier .xlsx cu mai multe sheet-uri (in functie de numărul de grupe alese).

# Convenții
Programul pleacă de la următoarele premise (țineți cont să fie respectate, pentru funcționarea corectă a programului):

- aveți instalat Python 3 si librăriile XLSXWriter și openpyxl pentru Python 3
- folosiți ca sursă un [orar](https://i.imgur.com/25szy1B.png) al [Universității Transilvania](http://www.unitbv.ro), fișier valid .xls/.xlsx, la care aveți drepturi de citire
- **numele fișierului sursă trebuie să se termine în "-Vx.xls" sau "-Vx.xlsx", deoarece programul preia versiunea orarului din titlu - un exemplu de nume corect este "Orar-semI-2017-2018-V13.xlsx"**
- **orarul fiecărei grupe trebuie să se întindă pe 4 linii, iar anul, specializarea și grupa trebuie să poată fi accesibile de pe un rând pe care se află și orarul (pe scurt, [anul, specializarea și grupa trebuie să se afle pe același rând cu orarul în sine](https://i.imgur.com/H9ZJMVu.png))**
- **codul orarului și anul universitar trebuie să fie câmpuri existente separate in orar** - locația se poate seta din script
- **fișierele "materii.txt", "materii_ignorate.txt" și "profesori.txt" trebuie să fie neapărat prezente lângă script, iar acestea trebuie să conțină o sintaxă corectă**
- setările din antetul codului sursă trebuie sa fie valide (tipul de date să fie cel potrivit)
- datele de intrare trebuie să fie valide - **PROGRAMUL NU VERIFICĂ CORECTITUDINEA DATELOR INTRODUSE**

# Instalare
Programul are nevoie de Python 3 și de librăriile **XLSXWriter** și **openpyxl** pentru Python 3. Nu voi acoperi instalarea Python 3, ci doar instalarea librăriilor necesare (asigurați-vă ca [aveți Python adăugat la PATH](https://i.imgur.com/QxIWjLX.png)). Pentru aceasta, rulați următoarele comenzi intr-un **Command Prompt cu drepturi de administrator**:

`pip install xlsxwriter openpyxl`

# Utilizare
Scriptul trebuie apelat dintr-un Command Prompt în felul următor (e recomandat, în cazul în care aveți spații în parametri, să puneți între ghilimele fiecare parametru):
`cd /d folder_in_care_se_afla_scriptul
python unitbv_generator_orar.py sursă folder_destinație rând_grupă_1 [rând_grupă_2 ...]`

- **sursă** - reprezintă calea către orarul UNITBV (ex. "D:\Downloads\Browser\Orar-semI-2017-2018-V13.xlsx")
- **folder_destinație** - reprezintă calea către folderul în care doriți să se genereze orarul pe grupe (ex. "D:\Downloads\Browser") - **programul va genera automat un fișier cu numele "orar-cod_orar.xlsx"**
- **rând_grupă_x** - reprezintă rândul de pe care începe orarul grupei dorite (ex. [în acest caz](https://i.imgur.com/ywVhiHd.png), pentru a genera orarul grupei 10LF271, vom introduce 16)

Exemplu de apel (va genera orarul pentru grupele prezente pe rândurile 16, 24 și 8):
`unitbv_generator_orar.py D:\Downloads\Orar-semI-2017-2018-V13.xlsx D:\Documents 16 24 8`

Orarul extras poate fi customizat folosind cele 3 fișiere care vin impreună cu scriptul:

- **materii.txt** - fișierul conține informații despre materiile din orar (numele cu care materia se găsește în orar, numele pe care materia il are în orarul generat și culoarea căsuței din orarul generat) - conținutul trebuie să urmeze sintaxa următoare (o linie separată pentru fiecare intrare):
`nume_materie_în_sursă=nume_materie_la_destinație=#culoare_la_destinație_în_hex`
- **materii_ignorate.txt** - fișierul conține numele din sursă ale materiilor pe care nu le doriți incluse în orarul generat
`nume_materie_în_sursă`
- **profesori.txt** - fișierul conține numele profesorului așa cum se găsește în orar și numele pe care îl va avea în orarul generat, separate prin "="
`nume_profesor_în_sursă=nume_profesor_la_destinație`

**În cazul în care nu există o intrare corespunzătoare în aceste fișiere, se vor folosi informațiile disponibile în sursă.**

# Setări (editabile în script)

- **__coord_cod_orar** - locația în spreadsheet a celulei care conține codul orarului (ex. [aici este "E1"](https://i.imgur.com/zwIZe4Q.png)
- **__coord_an_universitar** - locația în spreadsheet a celulei care conține anul universitar
- **__col_an** - coloana pe care se găsește anul de studiu
- **__col_spec** - coloana pe care se găsește specializarea
- **__col_grupa** - - coloana pe care se găsește grupa
- **__col_inceput_cursuri** - coloana pe care se găsește ora 8 în orarul grupei (ex. [aici](https://i.imgur.com/s1t7dRR.png), pentru 10LF271, coloana este E
- **__font** - numele fontului pe care il va folosi orarul generat
- **__marime_font** - mărimea fontului de mai sus
- **__culoare_border** - culoarea borderului în format hex

În cazul modificării formatului orarului, următoarele liste pot fi modificate corespunzător (se pot elimina elemente începand de la coadă, sau se pot adăuga):
- **__header_time** - conține headerele cu orele
- **__header_day** - conține headerele cu zilele

# Copyright
This software uses the XlsxWriter library, licensed under BSD license. Copyright (c) 2013, John McNamara <jmcnamara@cpan.org> All rights reserved.

This software uses the openpyxl library, licensed under MIT license. Copyright (c) 2017, Eric Gazoni, Charlie Clark.