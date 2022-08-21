import os, pdfplumber, re, time, csv
from natsort import os_sorted

def load_files(results):
    onlyfiles= []
    p = ""
    while not onlyfiles:
        while not os.path.exists(p):
            p: str = input("Podaj sciezke z plikami\n")
        os.chdir(p)
        while os.path.exists(results_file):
            try:
                os.remove(results_file)
            except:
                print(f"Zamknij proszę plik Excel z wynikami: {results}")
                input("Nacisnij Enter aby kontynuowac")
        onlyfiles = os_sorted([f for f in os.listdir(p) if os.path.isfile(os.path.join(p, f)) & (f.endswith(".txt") | f.endswith(".pdf"))])
        if not onlyfiles:
            print("Nie znaleziono w folderze plikow")
            p = "x"
    return onlyfiles

def file_to_text(filename):
    text = []
    soft = ""
    gear = ""
    hyperlinks:bool
    table_settings = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 1,
        "intersection_x_tolerance": 1, }
    if filename.split('.')[-1] == "txt":
        f = open(filename)
        text.append(f.read())
        f.close()
    elif filename.split('.')[-1] == "pdf":
        f = pdfplumber.open(filename)
        soft = f.metadata["Author"].lower()
        gear = f.metadata["Title"].lower()
        if "zeiss" in soft:
            if "gear" in gear:
                k = f.pages[0]
                cropped_gear = k.crop((0, 356.4, 612, 396)) # dla gear (0 * float(k.width), 0.45 * float(k.height), 1 * float(k.width), 0.5 * float(k.height))
                try:
                    text.append(re.sub("[a-z]", "", cropped_gear.extract_text()))
                except:
                    text.append(None)
                cropped_helix = k.crop((91.8, 696.96, 563, 752.4 ) )# dla helix (0.15 * float(k.width), 0.88 * float(k.height), 0.92 * float(k.width), 0.95 * float(k.height))
                try:
                    text.append(re.sub("[a-z]", "", cropped_helix.extract_text()))
                except:
                    text.append(None)
            else:
                if f.hyperlinks:
                    for k in f.pages:
                        text.append(k.extract_text(x_tolerance=3, y_tolerance=0))
                else:
                    for k in f.pages:
                        table = k.extract_table(table_settings)
                        if not table or "" in table[0]:
                            text.append(re.sub("\|-+|-+\|", "", k.extract_text(x_tolerance=3, y_tolerance=0)))
                            soft="old_zeiss"
        else:
            for k in f.pages:
                cropped = k.crop((0 * float(k.width), 0 * float(k.height), 0.5 * float(k.width), 0.95 * float(k.height)) )
                table = cropped.extract_table({"vertical_strategy": "lines",
        "horizontal_strategy": "text",
        "snap_tolerance": 20,
        "intersection_x_tolerance": 15,
        "join_tolerance": 15,
        "min_words_horizontal": 2,
        "keep_blank_chars": True})
                for j in table:
                    #extracted = (j[0] + "  " + j[1]) if j[1] is not None else j[0]
                    extracted = j[0]+ "\n" + j[1]
                    text.append(extracted)
    return text, gear, soft, select_pattern(soft, gear, filename.split('.')[-1])

def select_pattern(soft, gear, type):
    if type == "pdf":
        if "zeiss" in soft:
            ###if "gear" in gear or "helix" in gear:
                ###pattern = "(?: [-±>]?\d{1,3})(?: [-±>]?\d{1,3})( (?:[-±>]?)\d{1,3})( (?:[-±>]?)\d{1,3})( (?:[-±>]?)\d{1,3})( (?:[-±>]?)\d{1,3})" #pattern gear/helix
            if "old_zeiss" in soft:
                pattern = "(.+?)\n(?: +\d{1,4}[,.]\d{1,5}\n)?(?: +\n)? +(-?\d{1,4}[,.]\d{2,5})(?: +-?\d{1,4}[,.]\d{1,5}){2}" #pattern OLD Zeiss
            else:
                pattern = "(.+?) (-?\d{1,4}[,.]\d{1,5})(?: -?\d{1,4}[,.]\d{1,5}){2}"  # pattern NEW Zeiss
        else:
            pattern = re.compile("^(.+?)  (?:.+?)(\d{1,4}[.]\d{1,5})$", re.DOTALL)
    else:
        pattern = "(?:.+?);(.+?;.+?);(-?\d{1,4}[,.]\d{1,5});(?:-?\d{1,4}[,.]\d{1,5});(?:-?\d{1,4}[,.]\d{1,5})" #pattern txt Wenzel
    return pattern

def write_dict(d, tuples, count):
    for h in tuples:
        char_name = h[0].strip(' |-')
        d.setdefault(char_name, {})  #### dopisz do dict nazwa + wynik, usun |----
        d[char_name].setdefault(count, h[1].replace(".", ","))

def create_tuples(tuples, values, keys):
    for count, k in enumerate(keys):
        [tuples.append(j) for j in tuple(zip(k, values[count]))]

def from_dict_to_list(d):
    r = []
    if not d:
        return []
    m = max([max(val.keys()) for _, val in d.items()])
    for i in d.values():
        r.append([i[j] if j in i.keys() else "0" for j in range(m+1)])
    return r



if __name__ == '__main__':
    results = {}
    file_names = []
    time_of_measurement = []
    finded_tuples =[]
    results_file = "wyniki.csv"
    header =["Plik", "Data modyfikacji pliku (pomiaru)"]
    keys_gear = (('Falfa_lewa_1', 'Falfa_lewa_2', 'Falfa_lewa_3', 'Falfa_lewa_4'),
                 ('Falfa_prawa_1', 'Falfa_prawa_2', 'Falfa_prawa_3', 'Falfa_prawa_4'),
                 ('Ffalfa_lewa_1', 'Ffalfa_lewa_2', 'Ffalfa_lewa_3', 'Ffalfa_lewa_4'),
                 ('Ffalfa_prawa_1', 'Ffalfa_prawa_2', 'Ffalfa_prawa_3', 'Ffalfa_prawa_4'),
                 ('fHalfa_lewa_1', 'fHalfa_lewa_2', 'fHalfa_lewa_3', 'fHalfa_lewa_4'),
                 ('fHalfa_prawa_1', 'fHalfa_prawa_2', 'fHalfa_prawa_3', 'fHalfa_prawa_4'))
    keys_helix = (('FB_lewa_1', 'FB_lewa_2', 'FB_lewa_3', 'FB_lewa_4'),
                  ('FB_prawa_1', 'FB_prawa_2', 'FB_prawa_3', 'FB_prawa_4'),
                  ('FfB_lewa_1', 'FfB_lewa_2', 'FfB_lewa_3', 'FfB_lewa_4'),
                  ('FfB_prawa_1', 'FfB_prawa_2', 'FfB_prawa_3', 'FfB_prawa_4'),
                  ('fHB_lewa_1', 'fHB_lewa_2', 'fHB_lewa_3', 'fHB_lewa_4'),
                  ('fHB_prawa_1', 'fHB_prawa_2', 'fHB_prawa_3', 'fHB_prawa_4'))

    print("Skrypt wyszukuje wszystkie wartosci dla charakterystyk pomiarowych plikach TXT i PDF\n")
    print("Przetestowane na plikach z Wenzla i z Zeissa\n")

    files = load_files(results_file)
    print("Działam...")

    for count_files, i in enumerate(files):
        finded_tuples = []
        text, gear, soft, pattern = file_to_text(i)
        time_of_measurement.append(time.strftime("%d.%m.%Y;%H:%M", time.strptime(time.ctime(os.path.getmtime(i)), "%a %b  %d %H:%M:%S %Y")))
        file_names.append(i)
        if "gear" in gear:
            if text[0]:
            ##finded_values = re.findall(pattern, text[0])
            ###create_tuples(finded_tuples, finded_values, keys_gear)
                splitted = text[0].split()
                create_tuples(finded_tuples, (splitted[4:8], splitted[10:14], splitted[21:25], splitted[27:31], splitted[36:40], splitted[42:46]), keys_gear)
                write_dict(results, finded_tuples, count_files)
            if text[1]:
            ###finded_values = re.findall(pattern, text[0])
            ###create_tuples(finded_tuples, finded_values, keys_helix)
                splitted = text[1].split()
                create_tuples(finded_tuples, (splitted[3:7], splitted[9:13], splitted[18:22], splitted[24:27], splitted[33:37], splitted[39:43]), keys_helix)
                write_dict(results, finded_tuples, count_files)
        else:
            for j in text:
                finded_tuples = re.findall(pattern, j)
                write_dict(results, finded_tuples, count_files)
    header.extend(results.keys())

    print("Zapisuje plik...")
    with open(results_file, mode='w', newline='', encoding='utf-8') as file:
        wr = csv.writer(file, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        wr.writerow(header)
        p = zip(*from_dict_to_list(results))  #### transpozycja matrycy wynikow (poziom -> pion)
        for count, row in enumerate(p):
            d = list(row)
            d.insert(0, time_of_measurement[count])
            d.insert(0, file_names[count])
            wr.writerow(d)
    os.startfile(results_file)