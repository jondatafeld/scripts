#!/usr/bin/env python3
"""
Vérifie, pour une liste de codeap donnés, s'il existe des flux sur le cluster de
logs, et liste les datasets associés. Sortie : un fichier CSV (ouvrable dans Excel),
une ligne = un codeap.

AUCUNE dépendance externe (utilise uniquement la bibliothèque standard Python).

Entrées :
  --codeaps : fichier texte listant les codeap à vérifier (un par ligne)
              accepte 'ap12345', 'a12345', ou juste '12345'
  --indices : sortie de  GET _cat/indices/logs-*?h=index
  --output  : chemin du fichier .csv à produire (défaut: codeap_flux.csv)

Usage :
  python3 codeap_flux_csv.py --codeaps codeaps.txt --indices indices.txt --output resultat.csv
"""

import re
import csv
import argparse
from collections import defaultdict

FLUX_PATTERN = re.compile(
    r"^logs-"
    r"(?P<codename>[^.]+)\."
    r"(?P<dataset>.+)"
    r"-(?P<codeap>ap?\d+)"
    r"_(?P<env>[a-zA-Z0-9]+)"
    r"_(?P<retention>r\d+)"
    r"_(?P<date>.+)$"
)


def normalize_codeap(raw: str) -> str:
    """Ne garde que les chiffres pour matcher quel que soit le préfixe (a / ap / rien)."""
    return re.sub(r"\D", "", raw.strip())


def extract_index_name(line: str):
    for tok in line.split():
        if tok.startswith("logs-"):
            return tok
    return None


def main():
    parser = argparse.ArgumentParser(description="Vérifie la présence de flux par codeap (sortie CSV).")
    parser.add_argument("--codeaps", required=True, help="Fichier des codeap (un par ligne)")
    parser.add_argument("--indices", required=True, help="Fichier des noms d'index (_cat/indices)")
    parser.add_argument("--output", default="codeap_flux.csv", help="Fichier CSV de sortie")
    args = parser.parse_args()

    # Inventaire du cluster : codeap normalisé -> infos flux
    inventory = defaultdict(lambda: {
        "datasets": set(),
        "codenames": set(),
        "envs": set(),
        "index_count": 0,
    })

    with open(args.indices, "r", encoding="utf-8") as f:
        for line in f:
            idx = extract_index_name(line)
            if idx is None:
                continue
            m = FLUX_PATTERN.match(idx)
            if not m:
                continue
            key = normalize_codeap(m.group("codeap"))
            entry = inventory[key]
            entry["datasets"].add(m.group("dataset"))
            entry["codenames"].add(m.group("codename"))
            entry["envs"].add(m.group("env"))
            entry["index_count"] += 1

    # Liste des codeap à vérifier
    codeaps_input = []
    with open(args.codeaps, "r", encoding="utf-8") as f:
        for line in f:
            val = line.strip()
            if val:
                codeaps_input.append(val)

    # Écriture CSV — séparateur ';' pour ouverture directe dans Excel FR
    # (Excel français interprète ';' comme séparateur de colonnes par défaut)
    n_present = 0
    with open(args.output, "w", newline="", encoding="utf-8-sig") as f:
        # encoding utf-8-sig ajoute le BOM pour qu'Excel affiche correctement les accents
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Codeap", "Flux présent ?", "Nb datasets", "Datasets",
                         "Codenames", "Environnements", "Nb index"])

        for raw_codeap in codeaps_input:
            key = normalize_codeap(raw_codeap)
            entry = inventory.get(key)
            present = entry is not None and entry["index_count"] > 0

            if present:
                n_present += 1
                datasets = sorted(entry["datasets"])
                writer.writerow([
                    raw_codeap,
                    "OUI",
                    len(datasets),
                    ", ".join(datasets),
                    ", ".join(sorted(entry["codenames"])),
                    ", ".join(sorted(entry["envs"])),
                    entry["index_count"],
                ])
            else:
                writer.writerow([raw_codeap, "NON", 0, "", "", "", 0])

    print(f"Codeap vérifiés : {len(codeaps_input)}")
    print(f"  avec flux     : {n_present}")
    print(f"  sans flux     : {len(codeaps_input) - n_present}")
    print(f"Fichier généré  : {args.output}")


if __name__ == "__main__":
    main()
