#!/usr/bin/env python3
"""Codification automatique d'articles consommables (Excel) avec Ollama (qwen2.5:3b)."""
import csv
import json
import re
import sys
import time
import unicodedata
from copy import copy
from datetime import datetime
from pathlib import Path

import requests
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.formula.translate import Translator

# ----------------------------- Constantes ---------------------------------
MODELE = "qwen2.5:3b"
URL_OLLAMA = "http://localhost:11434"
BASE = Path(__file__).resolve().parent
DOSSIER_INPUT = BASE / "input"
DOSSIER_OUTPUT = BASE / "output"
JAUNE = "FFFFFF00"
LONGUEUR_MAX = 30
NUM_CTX = 4096          # contexte réduit = moins de RAM / CPU
MAX_EXEMPLES = 12       # exemples envoyés à l'étape 2 (strict nécessaire)
# Colonnes (numéros) : A=N°, B=NIMRODA, C=fournisseur, D=réf, I/J/L/O=formules
COL_K, COL_M, COL_N, COL_P, COL_Q = 11, 13, 14, 16, 17
COLS_FORMULES = (9, 10, 12, 15)

REGLES = """Format : FAMILLE-SOUSFAMILLE-DETAIL, 30 caracteres maximum, majuscules, sans espace ni accent.
Caracteres autorises dans le DETAIL : A-Z 0-9 . - +
Regles du DETAIL :
1. Un abrasif ou un foret porte toujours sa cote, jamais STD.
   - Diametre = D + valeur : D125, D50, D1.6
   - Deux dimensions separees par un point : 16.26, 23.28, D13.10
   - Grain = G + valeur : G120, G80, G36+ ; grains en mots : GMOY (moyen), GFIN (fin / tres fin / ultra fin)
   - Outils rotatifs sur tige (roues, manchons) : D<diam>.<largeur>-G<grain>, ex. D6.10-G150, D30.15-G60
   - Produits plats (bandes, feuilles, rouleaux) : <dim1>.<dim2>G<grain>, ex. 23.28G400, 300.25G40
2. Un support reprend la cote du consommable qu'il porte : support de manchon D6.10 <-> manchon D6.10-G150 ; porte-capuchon 16.26 <-> capuchon 16.26G120.
3. Taille = T + valeur : T39, T10, T2XL ; tailles en lettres : TXL, TL, TM, TS.
4. Volume : 5L ; si deux produits ont le meme volume dans la meme sous-famille, ajouter la reference produit : DR62-5L.
5. STD est reserve a un EPI ou un outil unique sans variante (casquette, cutter).
6. Si la designation NIMRODA et la designation fournisseur different (ex. 20.10 contre 20X15), la designation fournisseur l'emporte.
7. Le code complet doit etre unique dans tout le fichier."""


# ----------------------------- Utilitaires --------------------------------
def txt(v):
    """Valeur de cellule -> texte nettoyé ('' si vide)."""
    return "" if v is None else str(v).strip()


def est_jaune(cell):
    """Vrai si la cellule a un remplissage jaune FFFFFF00."""
    f = cell.fill
    return bool(f and f.fill_type == "solid" and f.fgColor.type == "rgb" and f.fgColor.rgb == JAUNE)


def nettoyer(detail):
    """Majuscules, sans accents ni espaces, uniquement A-Z 0-9 . - +"""
    s = unicodedata.normalize("NFKD", str(detail))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^A-Z0-9.+\-]", "", re.sub(r"\s+", "", s.upper()))


def verifier_ollama():
    """Arrête le programme si Ollama ou le modèle sont absents."""
    try:
        noms = [m["name"] for m in requests.get(f"{URL_OLLAMA}/api/tags", timeout=5).json()["models"]]
    except Exception:
        sys.exit("ERREUR : Ollama ne repond pas. Lance Ollama puis relance ce programme.")
    if not any(n == MODELE or n.startswith(MODELE + ":") for n in noms):
        sys.exit(f"ERREUR : modele absent. Tape dans un terminal :  ollama pull {MODELE}")


def appeler_ia(fixe, variable, titre):
    """Un appel Ollama indépendant (aucun historique). La partie fixe est placée en premier
    pour que Ollama réutilise son cache. Affiche entrée / sortie en temps réel."""
    print(f"     [{titre}] ENVOI ({len(fixe) + len(variable)} car.) :", flush=True)
    print("       " + variable.replace("\n", "\n       "), flush=True)
    debut = time.time()
    r = requests.post(f"{URL_OLLAMA}/api/chat", timeout=600, json={
        "model": MODELE, "stream": False, "format": "json", "keep_alive": "30m",
        "options": {"temperature": 0, "num_ctx": NUM_CTX},
        "messages": [{"role": "user", "content": fixe + "\n\n" + variable}]})
    r.raise_for_status()
    brut = r.json()["message"]["content"]
    print(f"     [{titre}] REPONSE en {time.time() - debut:.1f}s : {brut.strip()}", flush=True)
    try:
        rep = json.loads(brut)
        return rep if isinstance(rep, dict) else {}
    except ValueError:
        return {}


# ----------------------------- Lecture du fichier -------------------------
def lire_referentiel(ws):
    """Renvoie (familles {code: libellé}, sous-familles {clé: libellé}) lus dans l'onglet."""
    fam, sub = {}, {}
    for a, b, _, d, e in ws.iter_rows(min_row=2, max_col=5, values_only=True):
        if txt(a):
            fam[txt(a).upper()] = txt(b)
        if txt(d):
            sub[txt(d).upper()] = txt(e)
    return fam, sub


def ligne_info(ws, r):
    """Dictionnaire des champs utiles d'une ligne."""
    g = lambda c: txt(ws.cell(r, c).value)
    return {"ligne": r, "num": g(1), "nim": g(2), "four": g(3), "ref": g(4),
            "K": g(COL_K).upper(), "M": g(COL_M).upper(), "N": g(COL_N).upper()}


def code_de(i):
    return f"{i['K']}-{i['M']}-{i['N']}" if i["K"] and i["M"] and i["N"] else ""


def fournisseur_utile(i):
    return "" if i["four"].upper() == "X" else i["four"]


# ----------------------------- Les 2 étapes IA ----------------------------
def choisir_sous_famille(i, fam, sub, modeles):
    """Étape 1 : renvoie (clé, None) ou (None, famille_probable)."""
    print("  Etape 1/2 : choix de la sous-famille", flush=True)
    liste = "\n".join(f"{k} : {lib}" for k, lib in sub.items())
    ex = []
    for k in sub:  # 1 seul exemple par sous-famille
        for m in [m for m in modeles if f"{m['K']}-{m['M']}" == k][:1]:
            ex.append(f"{m['nim']} -> {k}")
    fixe = ("Choisis la cle de sous-famille de l'article. Reponds en JSON : {\"cle\": \"ABR-RL\"} "
            "ou {\"cle\": \"AUCUNE\"} si rien ne convient.\n\nCles possibles :\n" + liste +
            "\n\nExemples :\n" + "\n".join(ex))
    var = f"Article : {i['nim']}" + (f" | fournisseur : {fournisseur_utile(i)}" if fournisseur_utile(i) else "")
    rep = ""
    for essai in range(3):
        suite = "" if essai == 0 else "\nreponds uniquement par une cle de la liste."
        rep = txt(appeler_ia(fixe, var + suite, f"etape 1, essai {essai + 1}").get("cle")).upper()
        if rep in sub:
            return rep, None
    pref = rep.split("-")[0]
    return None, (pref if pref in fam else None)


def construire_detail(i, cle, modeles, utilises):
    """Étape 2 : renvoie (détail, problème) ; problème = None si tout est validé."""
    print(f"  Etape 2/2 : construction du detail ({cle})", flush=True)
    pareils = [m for m in modeles if f"{m['K']}-{m['M']}" == cle] \
        or [m for m in modeles if m["K"] == cle.split("-")[0]]
    exemples = "\n".join(f"{m['nim']} | {fournisseur_utile(m)} -> {m['N']}" for m in pareils[-MAX_EXEMPLES:])
    deja = ", ".join(sorted(c.split("-", 2)[2] for c in utilises if c.startswith(cle + "-"))[-30:])
    fixe = (f"{REGLES}\n\nSous-famille : {cle}\nExemples valides (NIMRODA | fournisseur -> DETAIL) :\n{exemples}\n"
            f"Details deja utilises : {deja or 'aucun'}\n\n"
            'Reponds en JSON : {"detail": "D30.15-G60"} (uniquement le DETAIL).')
    var = f"Article : {i['nim']} | fournisseur : {fournisseur_utile(i) or 'inconnue'}"
    detail, probleme = "", "detail vide"
    for essai in range(2):
        suite = "" if essai == 0 else f"\nATTENTION : {probleme} pour '{detail}'. Propose un autre detail."
        detail = nettoyer(appeler_ia(fixe, var + suite, f"etape 2, essai {essai + 1}").get("detail", ""))
        if detail.startswith(cle + "-"):
            detail = detail[len(cle) + 1:]
        if not detail:
            probleme = "detail vide"
        elif len(f"{cle}-{detail}") > LONGUEUR_MAX:
            probleme = f"code trop long (max {LONGUEUR_MAX})"
        elif f"{cle}-{detail}" in utilises:
            probleme = "collision : code deja utilise"
        else:
            return detail, None
    return detail, probleme


# ----------------------------- Réparation des lignes X --------------------
def reparer_ligne(ws, r, modele):
    """Recopie formules (I,J,L,O) et style (B..Q) d'une ligne modèle vers la ligne r."""
    for c in COLS_FORMULES:
        f = ws.cell(modele, c).value
        if not str(ws.cell(r, c).value or "").startswith("="):
            ws.cell(r, c).value = Translator(f, origin=ws.cell(modele, c).coordinate) \
                .translate_formula(ws.cell(r, c).coordinate)
    for c in range(2, 18):
        ws.cell(r, c)._style = copy(ws.cell(modele, c)._style)


# ----------------------------- Traitement d'un fichier --------------------
def traiter_fichier(chemin, horodatage, rapport):
    wb = load_workbook(chemin)  # sans data_only : on garde toutes les formules
    if "Codification" not in wb.sheetnames or "Referentiel" not in wb.sheetnames:
        print(f"  -> onglets 'Codification' / 'Referentiel' introuvables, fichier ignore.")
        return
    ws, fam, sub = wb["Codification"], *lire_referentiel(wb["Referentiel"])
    infos = [ligne_info(ws, r) for r in range(2, ws.max_row + 1) if txt(ws.cell(r, 2).value)]
    a_traiter = [i for i in infos if est_jaune(ws.cell(i["ligne"], 1)) or not code_de(i)]
    numeros = {i["ligne"] for i in a_traiter}
    modeles = [i for i in infos if i["ligne"] not in numeros]
    utilises = {code_de(m) for m in modeles}
    tpl = next((m["ligne"] for m in modeles if all(
        str(ws.cell(m["ligne"], c).value or "").startswith("=") for c in COLS_FORMULES)), None)
    print(f"  {len(modeles)} lignes modeles, {len(a_traiter)} lignes a traiter.")
    stats = {"ok": 0, "verif": 0}

    for i in a_traiter:
        r, ancien = i["ligne"], code_de(i) or "aucun"
        print(f"\n>>> Ligne {r}/{ws.max_row} : {i['nim']}", flush=True)
        K = M = N = None
        try:
            cle, fam_probable = choisir_sous_famille(i, fam, sub, modeles)
            if cle:
                K, M = cle.split("-", 1)
                N, probleme = construire_detail(i, cle, modeles, utilises)
                statut = "OK" if not probleme else f"A VERIFIER : {probleme.split(' (')[0].split(' :')[0]}"
                if not N:
                    N = None
                else:
                    utilises.add(f"{cle}-{N}")
                if not probleme and not fournisseur_utile(i):
                    statut = "OK (sans designation fournisseur)"
            else:
                K, statut = fam_probable, "A VERIFIER : sous-famille introuvable"
        except Exception as e:  # toute erreur : on marque la ligne et on continue
            print(f"     erreur ligne {r} : {e}")
            statut = "A VERIFIER : erreur IA"
        if tpl and not all(str(ws.cell(r, c).value or "").startswith("=") for c in COLS_FORMULES):
            reparer_ligne(ws, r, tpl)
        ws.cell(r, COL_K).value, ws.cell(r, COL_M).value, ws.cell(r, COL_N).value = K, M, N
        ws.cell(r, COL_P).value = statut
        ws.cell(r, COL_N).comment = Comment(f"Ancien code : {ancien}", "Codification IA")
        nouveau = f"{K}-{M}-{N}" if K and M and N else ""
        stats["ok" if statut.startswith("OK") else "verif"] += 1
        print(f"Ligne {r}/{ws.max_row} : {nouveau or statut}")
        rapport.append([chemin.name, r, i["num"], i["nim"], ancien, nouveau, statut])

    # Collisions : doublons du code complet sur TOUT le fichier
    groupes = {}
    for r in range(2, ws.max_row + 1):
        if txt(ws.cell(r, 2).value):
            i = ligne_info(ws, r)
            ws.cell(r, COL_Q).value = None
            if code_de(i):
                groupes.setdefault(code_de(i), []).append(r)
    collisions = 0
    for lignes in groupes.values():
        if len(lignes) > 1:
            collisions += len(lignes)
            for r in lignes:
                ws.cell(r, COL_Q).value = f"OUI (lignes {', '.join(map(str, lignes))})"

    sortie = DOSSIER_OUTPUT / f"{chemin.stem}_codifie_{horodatage}.xlsx"
    wb.save(sortie)
    print(f"\nResume : {len(a_traiter)} traitees | {stats['ok']} OK | {stats['verif']} a verifier | "
          f"{collisions} lignes en collision\nFichier : {sortie}\n")


def main():
    DOSSIER_INPUT.mkdir(exist_ok=True)
    DOSSIER_OUTPUT.mkdir(exist_ok=True)
    fichiers = sorted(p for p in DOSSIER_INPUT.glob("*.xlsx") if not p.name.startswith("~$"))
    if not fichiers:
        sys.exit(f"Le dossier '{DOSSIER_INPUT.name}' ne contient aucun fichier .xlsx : depose ton fichier Excel dedans.")
    verifier_ollama()
    maintenant = datetime.now()
    horodatage = maintenant.strftime("%Y-%m-%d_%Hh%M")
    rapport = []
    for f in fichiers:
        print(f"\n=== {f.name} ===")
        traiter_fichier(f, horodatage, rapport)
    csv_path = DOSSIER_OUTPUT / f"rapport_{maintenant.strftime('%Y-%m-%d')}.csv"
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["fichier", "ligne", "N°", "designation NIMRODA", "ancien code", "nouveau code", "statut"])
        w.writerows(rapport)
    print(f"Rapport : {csv_path}")


if __name__ == "__main__":
    main()
