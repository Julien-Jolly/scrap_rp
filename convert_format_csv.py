import csv


def format_float(value):
    try:
        float_value = float(value)
        return f"{float_value:.2f}".replace(".", ",")
    except ValueError:
        return value


def convert_csv(path, mois_str, year):
    file_in = f"{path}booking_{mois_str}-{year}.csv"
    file_out = f"{path}booking_{mois_str}-{year}_converted.csv"

    with open(file_in, mode="r", newline="", encoding="utf-8") as file_in, open(
        file_out, mode="w", newline="", encoding="utf-8"
    ) as file_out:
        reader = csv.reader(file_in, delimiter=",")
        writer = csv.writer(file_out, delimiter=";")

        for l in reader:
            formated_line = [format_float(c) for c in l]
            writer.writerow(formated_line)

    print(f"Le fichier modifié a été enregistré sous '{file_out}'.")
