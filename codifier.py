#!/usr/bin/env python3
"""Codification automatique d'articles consommables (Excel), 100 % locale (Ollama).
Sous-famille : un modele d'embeddings compare le SENS de l'article aux libelles du Referentiel et aux lignes deja
validees ; si c'est net on decide, sinon le LLM tranche entre quelques candidates. Detail : regles lues dans l'Excel."""
import csv
import hashlib
import json
import math
import os
import re
import sys
import time
import unicodedata
from collections import Counter
from copy import copy
from datetime import datetime
from pathlib import Path

import requests
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.formula.translate import Translator
from openpyxl.styles import PatternFill

# ----------------------------- Constantes ---------------------------------
MODELE = "qwen3:4b"                 # LLM (repli plus leger : "qwen2.5:3b")
MODELE_EMB = "qwen3-embedding:0.6b"  # embeddings (comprend le francais)
URL_OLLAMA = "http://localhost:11434"
BASE = Path(__file__).resolve().parent
DOSSIER_INPUT, DOSSIER_OUTPUT = BASE / "input", BASE / "output"
JAUNE = "FFFF00"        # marque "a traiter" posee par l'utilisateur sur la colonne A
LONGUEUR_MAX = 30
NUM_CTX, NUM_PREDICT = 2048, 24               # contexte minimal = RAM / CPU legers
NUM_THREAD = max(2, (os.cpu_count() or 4) // 2)  # coeurs physiques (hyper-threading inutile)
ECART_DIRECT, ACCORD_DIRECT = 0.08, 3           # decision sans LLM : ecart de similarite, ou accord des voisins
K_VOISINS, NB_CANDIDATES = 6, 6
LIGNES_REFERENTIEL = 500  # les plages VLOOKUP vers le Referentiel sont elargies jusque-la
# role -> (regex cherchee dans l'en-tete de la ligne 1, colonne par defaut)
ROLES = {"K": ("code famille", 11), "M": ("code sous", 13), "N": ("^detail", 14),
         "statut": ("statut", 16), "collision": ("collision", 17),
         "nim": ("nimroda", 2), "four": ("designation fournisseur", 3), "ref": ("^ref", 4)}
PALETTE = {"jaune": "FFFF00", "rouge": "FF0000", "rouge clair": "F8D7DA", "bleu clair": "E8F0FE",
           "bleu fonce": "1F4E78", "orange clair": "FFD8A8", "jaune clair": "FFF3CD",
           "vert clair": "C6EFCE", "gris": "D9D9D9", "blanc": "FFFFFF"}
ROUGE, VERT, GRIS, FIN = "\x1b[31m", "\x1b[32m", "\x1b[90m", "\x1b[0m"


# ----------------------------- Utilitaires --------------------------------
def txt(v):
    return "" if v is None else str(v).strip()


def norm(t):
    """Minuscules sans accents."""
    s = unicodedata.normalize("NFKD", txt(t).lower())
    return "".join(c for c in s if not unicodedata.combining(c))


def nettoyer(detail):
    """Normalise un detail : 10*330 / 10X330 -> 10.330, 3,7 -> 3.7, majuscules, A-Z 0-9 . - +"""
    s = re.sub(r"(?<=\d)[*xX](?=\d)", ".", norm(detail))
    s = re.sub(r"(?<=\d),(?=\d)", ".", s)
    return re.sub(r"[^A-Z0-9.+\-]", "", s.upper())


def couleur(cell):
    """Couleur de fond RGB (6 hexa) ou None."""
    f = cell.fill
    if f and f.fill_type == "solid" and f.fgColor.type == "rgb":
        return str(f.fgColor.rgb)[-6:].upper()
    return None


def nom_couleur(rgb):
    p = [int(rgb[i:i + 2], 16) for i in (0, 2, 4)]
    return min(PALETTE, key=lambda n: sum((int(PALETTE[n][i:i + 2], 16) - p[j]) ** 2 for j, i in enumerate((0, 2, 4))))


def pastille(rgb):
    return f"\x1b[48;2;{int(rgb[0:2], 16)};{int(rgb[2:4], 16)};{int(rgb[4:6], 16)}m  {FIN}"


def verifier_ollama():
    """Arrete le programme si Ollama ou un modele est absent, puis charge le LLM en memoire."""
    try:
        noms = [m["name"] for m in requests.get(f"{URL_OLLAMA}/api/tags", timeout=5).json()["models"]]
    except Exception:
        sys.exit("ERREUR : Ollama ne repond pas. Lance Ollama puis relance ce programme.")
    for m in (MODELE, MODELE_EMB):
        if not any(n == m or n.startswith(m + ":") for n in noms):
            sys.exit(f"ERREUR : modele absent. Tape dans un terminal :  ollama pull {m}")
    print("Chargement du modele en memoire...", flush=True)
    appeler_ia("ok", "ok", "chargement", ["ok"])


def appeler_ia(fixe, variable, titre, choix):
    """Appel LLM independant (aucun historique). La reponse est contrainte a {"cle": <une valeur de choix>}."""
    print(f"     {GRIS}[{titre}] ENVOI ({len(fixe) + len(variable)} car.){FIN} " + variable.replace("\n", " / "), flush=True)
    debut = time.time()
    r = requests.post(f"{URL_OLLAMA}/api/chat", timeout=600, json={
        "model": MODELE, "stream": False, "think": False, "keep_alive": "30m",
        "format": {"type": "object", "properties": {"cle": {"type": "string", "enum": choix}}, "required": ["cle"]},
        "options": {"temperature": 0, "num_ctx": NUM_CTX, "num_predict": NUM_PREDICT, "num_thread": NUM_THREAD},
        "messages": [{"role": "user", "content": fixe + "\n\n" + variable}]})
    r.raise_for_status()
    brut = r.json()["message"]["content"]
    print(f"     {GRIS}[{titre}] REPONSE en {time.time() - debut:.1f}s :{FIN} {brut.strip()}", flush=True)
    try:
        return json.loads(brut)
    except ValueError:
        return {}


def embeddings(textes, cache):
    """Vecteurs de sens (modele d'embeddings) ; ce qui est deja dans le cache n'est pas recalcule."""
    cle = lambda t: hashlib.md5((MODELE_EMB + t).encode()).hexdigest()
    manque = [t for t in dict.fromkeys(textes) if cle(t) not in cache]
    for k in range(0, len(manque), 32):
        print(f"  Embeddings : {min(k + 32, len(manque))}/{len(manque)} textes", flush=True)
        r = requests.post(f"{URL_OLLAMA}/api/embed", timeout=600,
                          json={"model": MODELE_EMB, "input": manque[k:k + 32], "keep_alive": "30m"})
        r.raise_for_status()
        for t, v in zip(manque[k:k + 32], r.json()["embeddings"]):
            cache[cle(t)] = [round(x, 5) for x in v]
    return [cache[cle(t)] for t in textes]


def cos(a, b):
    return sum(x * y for x, y in zip(a, b)) / (math.sqrt(sum(x * x for x in a)) * math.sqrt(sum(y * y for y in b)))


# ----------------------------- Lecture du fichier -------------------------
def lire_legende(wb):
    """Onglet Legende : cellule coloree vide + texte a droite. Renvoie {rgb: sens} et les couleurs par role."""
    sens = {}
    if "Legende" in wb.sheetnames:
        for row in wb["Legende"].iter_rows():
            for c in row:
                rgb = couleur(c)
                if rgb and not txt(c.value) and txt(c.offset(0, 1).value):
                    sens[rgb] = txt(c.offset(0, 1).value)
    roles = {"edit": ("modifier", "a la main"), "collision": ("collision",),
             "traiter": ("a traiter", "a codifier", "a completer")}
    return sens, {r: {k for k, v in sens.items() if any(m in norm(v) for m in ms)} for r, ms in roles.items()}


def lire_referentiel(ws):
    """(familles {code: libelle}, sous-familles {cle: libelle}) lus dans le fichier."""
    fam, sub = {}, {}
    for a, b, _, d, e in ws.iter_rows(min_row=2, max_col=5, values_only=True):
        if txt(a):
            fam[txt(a).upper()] = txt(b)
        if txt(d):
            sub[txt(d).upper()] = txt(e)
    return fam, sub


def trouver_colonnes(ws, couleurs_edit):
    """Colonnes par en-tete ; a defaut, les 3 colonnes designees 'a modifier' par la legende."""
    col = {}
    for role, (mot, defaut) in ROLES.items():
        col[role] = next((c for c in range(1, ws.max_column + 1) if re.search(mot, norm(ws.cell(1, c).value))), defaut)
    editables = [c for c in range(1, ws.max_column + 1)
                 if any(couleur(ws.cell(r, c)) in couleurs_edit for r in range(2, min(ws.max_row, 40) + 1))]
    if len(editables) == 3 and not all(norm(ws.cell(1, c).value) for c in editables):
        col["K"], col["M"], col["N"] = editables
    return col, editables


def ligne_info(ws, r, col):
    g = lambda c: txt(ws.cell(r, c).value)
    return {"ligne": r, "num": g(1), "nim": g(col["nim"]), "four": g(col["four"]), "ref": g(col["ref"]),
            "K": g(col["K"]).upper(), "M": g(col["M"]).upper(), "N": g(col["N"]).upper()}


def code_de(i):
    return f"{i['K']}-{i['M']}-{i['N']}" if i["K"] and i["M"] and i["N"] else ""


def texte_article(i):
    return i["nim"] + ("" if i["four"].upper() == "X" else " " + i["four"])


# ----------------------------- Etape 1 : sous-famille ---------------------
def choisir_sous_famille(i, vi, vec_lib, voisins, sub):
    """Renvoie (cle, sure, autres candidates). Signaux : (1) similarite de sens avec le libelle de chaque sous-famille (marche meme
    sans exemple) ; (2) vote des lignes validees les plus proches. Net -> decision directe ; sinon le LLM tranche
    entre quelques candidates, et le choix n'est sur que s'il confirme l'un des deux signaux."""
    lib = sorted(((cos(vi, v), k) for k, v in vec_lib.items()), reverse=True)
    nb = sorted(((cos(vi, v), m) for m, v in voisins), key=lambda x: -x[0])[:K_VOISINS]
    vote = Counter()
    for s, m in nb[:5]:
        vote[f"{m['K']}-{m['M']}"] += s
    knn = vote.most_common(1)[0][0] if vote else None
    accord = sum(1 for s, m in nb[:5] if f"{m['K']}-{m['M']}" == knn)
    ecart, l1 = lib[0][0] - lib[1][0], lib[0][1]
    if ecart >= ECART_DIRECT or (l1 == knn and accord >= ACCORD_DIRECT):
        cle = l1 if ecart >= ECART_DIRECT else knn
        print(f"  Etape 1/2 : {VERT}{cle} (sens proche du libelle, ecart {ecart:.2f} ; voisins d'accord {accord}/5, 0 appel LLM){FIN}", flush=True)
        return cle, True, [k for s, k in lib[1:3]] + ([knn] if knn and knn != cle else [])
    cand = list(dict.fromkeys([k for s, k in lib[:3]] + [k for k, _ in vote.most_common(2)] + ([knn] if knn else [])))[:NB_CANDIDATES]
    print(f"  Etape 1/2 : ambigu (libelle {l1} ecart {ecart:.2f}, voisins {knn} {accord}/5) -> LLM sur {len(cand)} candidates", flush=True)
    fixe = ("Choisis la sous-famille de l'article de consommables industriels. Reponds en JSON {\"cle\": \"...\"}.\n"
            "Sous-familles possibles :\n" + "\n".join(f"{k} : {sub[k]}" for k in cand)
            + "\n\nExemples deja codes :\n" + "\n".join(f"{m['nim'][:60]} -> {m['K']}-{m['M']}" for s, m in nb))
    cle = txt(appeler_ia(fixe, f"Article : {texte_article(i)[:200]}", "etape 1", cand).get("cle"))
    cle = cle if cle in cand else l1
    return cle, cle in (l1, knn), [k for k in cand if k != cle]


# ----------------------------- Etape 2 : detail par regles ----------------
# Regles deduites de l'Excel (section 5) : chaque TYPE de detail est un format applique a des attributs lus
# dans les designations (fournisseur d'abord). Le type est fixe par sous-famille, jamais par ligne.
N = r"(\d+(?:[.,]\d+)?)"
KX = r"\s*[x*\u00d7]\s*"
TYPES = {  # type -> (format, description donnee a l'IA quand aucune ligne validee n'existe)
    "DIAM": (lambda a: (a["diam"] or a["dims"] and a["dims"][0]) and "D" + (a["diam"] or a["dims"][0]),
             "diametre seul : D6.5 (forets, brosses)"),
    "TAILLE": (lambda a: a["taille"] and "T" + a["taille"], "taille : T10 (gants, chaussures, vetements)"),
    "TIGE": (lambda a: a["dims"] and a["g"] and f"D{a['dims'][0]}.{a['dims'][1]}-G{a['g']}",
             "outil rotatif sur tige (roue, manchon) : D30.15-G80"),
    "DISQUE": (lambda a: a["diam"] and a["g"] and f"D{a['diam']}G{a['g']}", "disque : diametre et grain, D125G36"),
    "PLAT": (lambda a: a["dims"] and a["g"] and f"{a['dims'][0]}.{a['dims'][1]}G{a['g']}",
             "produit plat (bande, feuille, capuchon) : dim1.dim2 et grain, 23.28G400"),
    "GRAIN": (lambda a: a["g"] and "G" + a["g"], "grain seul : G80 (rouleaux, disques de poncage)"),
    "COTE": (lambda a: a["dims"] and f"D{a['dims'][0]}.{a['dims'][1]}", "support : cotes du consommable porte, D6.10"),
    "VOLUME": (lambda a: a["vol"] and a["vol"] + "L", "produit liquide : volume, 5L"),
    "LARGEUR": (lambda a: a["mm"] and "L" + a["mm"], "ruban, rouleau adhesif : L25 (largeur en mm)"),
    "LUMENS": (lambda a: a["lum"] and "L" + a["lum"], "lampe : L300 (lumens)"),
    "STD": (lambda a: "STD", "outil ou EPI unique sans variante : STD"),
}


def attributs(t):
    """Attributs lus dans UNE designation (deja normalisee par norm)."""
    m = lambda p: re.search(p, t)
    diam = next((_f(x.group(1)) for x in (m(r"(?:diametre|diam|dia)\.?\s*:?\s*" + N), m(r"\u00f8\s*" + N),
                                          m(r"\bd\.?\s*" + N + r"\s*mm"), m(N + r"\s*mm"), m(r"\bd\.?\s*" + N)) if x), None)
    dims = m(N + KX + N + "(?:" + KX + N + ")?") or m(r"\bd\s*" + N + r"\s*l\s*" + N) or m(r"\b(\d{1,3})\.(\d{1,3})\b(?!\s*mm)")
    g = m(r"\b(?:grain|gr|g|p)\.?\s*[:.]?\s*(?:a|ceramique|oxy\w*|alu\w*)?\s*" + N + r"\s*(\+)?")
    mot = "FIN" if m(r"\b(ultra |tres |very )?fin(e)?\b") else "MOY" if m(r"moyen|medium") else None
    ta = m(r"(?:taille|pointure)\s*[.:]?\s*" + N + r"(?:/(\d{1,2}(?:[.,]\d)?)(?!\d))?") or m(r"\bt\s*(\d{1,2})\b")
    vol, mm, lum = m(N + r"\s*(?:l|litres?)\b"), m(N + r"\s*mm"), m(N + r"\s*lumens?")
    dims = [_f(x) for x in dims.groups() if x] if dims else None
    return {"mm": mm and _f(mm.group(1)), "lum": lum and _f(lum.group(1)), "diam": diam, "dims": dims if dims and all(float(x) <= 999 for x in dims) else None,  # 100x10000 n'est pas une roue
            "g": mot or (_f(g.group(1)) + (g.group(2) or "") if g else None),
            "taille": ta and "".join(_f(x) for x in ta.groups() if x), "vol": vol and _f(vol.group(1))}


def _f(x):
    return x.replace(",", ".")


def extraire(i):
    """Attributs de l'article : designation fournisseur d'abord, NIMRODA pour completer ; '+' du grain si l'une le porte."""
    a = {}
    for t in (norm(i["four"]) if i["four"].upper() != "X" else "", norm(i["nim"])):
        for k, v in attributs(t).items():
            if v and not a.get(k):
                a[k] = v
        if re.search(r"\d\s*\+", t) and a.get("g", "").isdigit():
            a["g"] += "+"
    return {k: a.get(k) for k in ("diam", "dims", "g", "taille", "vol", "mm", "lum")}


def apprendre_types(lignes, seuil):
    """{cle: type} : pour chaque sous-famille, le type dont le format reproduit les codes des lignes
    (None si aucun ne les reproduit : codes propres au produit)."""
    par = {}
    for m in lignes:
        if code_de(m):
            par.setdefault(f"{m['K']}-{m['M']}", []).append(m)
    out = {}
    for cle, ex in par.items():
        taux = {t: sum(f(extraire(m)) == nettoyer(m["N"]) for m in ex) / len(ex) for t, (f, _) in TYPES.items()}
        t = max(taux, key=taux.get)
        out[cle] = t if taux[t] >= seuil else None
    return out


def type_par_ia(cle, sub, i, types, vec_lib):
    """Sous-famille sans aucun code connu : le LLM choisit le type par analogie avec les 6 sous-familles
    deja apprises dont le libelle est le plus proche (par le sens)."""
    proches = sorted((k for k, t in types.items() if t and k in vec_lib), key=lambda k: -cos(vec_lib[cle], vec_lib[k]))[:6]
    fixe = ("Quel type de detail convient a la sous-famille ? Reponds en JSON {\"cle\": \"TYPE\"}.\nTypes :\n"
            + "\n".join(f"{t} : {d}" for t, (_, d) in TYPES.items()) + "\n\nDeja etablis :\n"
            + "\n".join(f"- {sub[k]} -> {types[k]}" for k in proches))
    return txt(appeler_ia(fixe, f"Sous-famille : {sub[cle]}\nArticle : {texte_article(i)[:160]}", "type de detail", list(TYPES)).get("cle"))


def construire_detail(i, cle, typ, utilises):
    """Renvoie (detail, probleme) ; probleme = None si le code est valide et unique."""
    d = TYPES[typ][0](extraire(i)) if typ else None
    ref = re.search(r"\b([A-Z]{1,4})-?(\d{1,4}[A-Z]?)\b", i["nim"])  # regle 4 : DR-62 -> DR62-5L
    if d and typ == "VOLUME" and f"{cle}-{d}" in utilises and ref:
        d = ref.group(1) + ref.group(2) + "-" + d
    probleme = ("detail propre au produit" if not typ else "attribut introuvable" if not d else
                "code trop long" if len(f"{cle}-{d}") > LONGUEUR_MAX else "collision" if f"{cle}-{d}" in utilises else None)
    print(f"  Etape 2/2 : {VERT if not probleme else ROUGE}detail {d or '?'} (type {typ}){FIN}", flush=True)
    return d or None, probleme


# ----------------------------- Ecriture dans le classeur ------------------
def elargir_plages(ws):
    """Plages figees vers le Referentiel (ex. $A$2:$B$7) -> elargies, pour que tout ajout soit reconnu."""
    n = 0
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("=") and "Referentiel!" in c.value:
                new = re.sub(r"(Referentiel!\$[A-Z]+\$\d+:\$[A-Z]+\$)(\d+)",
                             lambda m: m.group(1) + str(max(int(m.group(2)), LIGNES_REFERENTIEL)), c.value)
                if new != c.value:
                    c.value, n = new, n + 1
    return n


def reparer_ligne(ws, r, modele, cols_formules):
    """Ligne X sans formules : recopie formules et style d'une ligne modele (retire le fond rouge)."""
    for c in cols_formules:
        if not txt(ws.cell(r, c).value).startswith("="):
            ws.cell(r, c).value = Translator(ws.cell(modele, c).value, origin=ws.cell(modele, c).coordinate) \
                .translate_formula(ws.cell(r, c).coordinate)
    for c in range(2, ws.max_column + 1):
        ws.cell(r, c)._style = copy(ws.cell(modele, c)._style)


def afficher_ligne(ws, r, sens):
    """Affiche la ligne entiere (valeurs de A a la derniere colonne) et ses couleurs."""
    print(f"\n>>> Ligne {r}/{ws.max_row}")
    champs, couleurs = [], {}
    for c in range(1, ws.max_column + 1):
        cell, v = ws.cell(r, c), txt(ws.cell(r, c).value)
        if v and not v.startswith("="):
            champs.append(f"{txt(ws.cell(1, c).value)[:22]}: {v[:80]}")
        if couleur(cell):
            couleurs.setdefault(couleur(cell), []).append(cell.column_letter)
    print("    " + " | ".join(champs))
    for rgb, lettres in couleurs.items():
        s = sens.get(rgb) or ("marque 'a traiter'" if rgb == JAUNE else "")
        print(f"    {pastille(rgb)} {nom_couleur(rgb)} (#{rgb}) colonnes {','.join(lettres)}" + (f" = {s}" if s else ""))


def traiter_fichier(chemin, horodatage, rapport, cache):
    wb = load_workbook(chemin)  # sans data_only : on garde toutes les formules
    if "Codification" not in wb.sheetnames or "Referentiel" not in wb.sheetnames:
        print("  -> onglets 'Codification' / 'Referentiel' introuvables, fichier ignore.")
        return
    ws, (fam, sub) = wb["Codification"], lire_referentiel(wb["Referentiel"])
    sens, roles = lire_legende(wb)
    col, editables = trouver_colonnes(ws, roles["edit"])
    lettre = lambda c: ws.cell(1, c).column_letter
    print(f"  Legende : {len(sens)} couleurs comprises ; cellules a modifier = colonnes "
          f"{', '.join(lettre(c) for c in editables) or '(non trouvees)'} ; utilise : "
          f"{lettre(col['K'])}, {lettre(col['M'])}, {lettre(col['N'])} (+ Statut {lettre(col['statut'])}, Collision {lettre(col['collision'])})")
    print(f"  Referentiel : {len(fam)} familles, {len(sub)} sous-familles ; {elargir_plages(ws)} formule(s) VLOOKUP elargie(s).")
    marques = {JAUNE} | roles["traiter"]
    infos = [ligne_info(ws, r, col) for r in range(2, ws.max_row + 1) if txt(ws.cell(r, col["nim"]).value)]
    a_traiter = [i for i in infos if couleur(ws.cell(i["ligne"], 1)) in marques or not code_de(i)]
    lignes_t = {i["ligne"] for i in a_traiter}
    modeles = [i for i in infos if i["ligne"] not in lignes_t]
    nb_f = lambda r: sum(txt(ws.cell(r, c).value).startswith("=") for c in range(1, ws.max_column + 1))
    tpl = max((m["ligne"] for m in modeles), key=nb_f, default=None)
    cols_f = [c for c in range(1, ws.max_column + 1) if tpl and txt(ws.cell(tpl, c).value).startswith("=")]
    utilises, par_ia = {code_de(m) for m in modeles}, set()
    types = {**apprendre_types(a_traiter, 0.8), **apprendre_types(modeles, 0.5)}  # lignes validees d'abord
    print(f"  Formats de detail appris : {sum(1 for t in types.values() if t)} sous-familles ; "
          f"{sum(1 for t in types.values() if not t)} a codes propres au produit.")
    print(f"  {len(modeles)} lignes modeles, {len(a_traiter)} lignes a traiter.")
    stats, debut = Counter(), time.time()
    vecs = embeddings([texte_article(m) for m in modeles] + [texte_article(i) for i in a_traiter] + list(sub.values()), cache)
    voisins, vec_art = list(zip(modeles, vecs)), vecs[len(modeles):len(modeles) + len(a_traiter)]
    vec_lib = dict(zip(sub, vecs[len(modeles) + len(a_traiter):]))

    for n, (i, vi) in enumerate(zip(a_traiter, vec_art)):
        r, ancien = i["ligne"], code_de(i) or "aucun"
        afficher_ligne(ws, r, sens)
        K = M = N = None
        try:
            cle, sure_sf, autres = choisir_sous_famille(i, vi, vec_lib, voisins, sub)
            for essai, c in enumerate([cle] + [a for a in autres if types.get(a)]):
                if c not in types:
                    types[c] = type_par_ia(c, sub, i, types, vec_lib)
                    par_ia.add(c)
                N, probleme = construire_detail(i, c, types[c], utilises)
                if N or not types[c]:  # trouve, ou codes propres au produit : on garde cette sous-famille
                    break
                print(f"     {GRIS}format de {c} inapplicable a cet article : on essaie la candidate suivante{FIN}", flush=True)
            if N and c != cle:
                cle, sure_sf = c, False  # sous-famille deduite du format : a confirmer
            K, M = cle.split("-", 1)
            statut = f"A VERIFIER : {probleme}" if probleme else "A VERIFIER : sous-famille incertaine" if not sure_sf \
                else "A VERIFIER : type de detail choisi par l'IA" if cle in par_ia else "OK"
            if N:
                utilises.add(f"{cle}-{N}")
            if statut == "OK" and i["four"].upper() in ("", "X"):
                statut = "OK (sans designation fournisseur)"
        except Exception as e:  # toute erreur : on marque la ligne et on continue
            print(f"     {ROUGE}erreur : {e}{FIN}")
            statut = "A VERIFIER : erreur IA"
        if tpl and cols_f and not all(txt(ws.cell(r, c).value).startswith("=") for c in cols_f):
            reparer_ligne(ws, r, tpl, cols_f)
        ws.cell(r, col["K"]).value, ws.cell(r, col["M"]).value, ws.cell(r, col["N"]).value = K, M, N or None
        ws.cell(r, col["statut"]).value = statut
        ws.cell(r, col["N"]).comment = Comment(f"Ancien code : {ancien}", "Codification IA")
        nouveau = f"{K}-{M}-{N}" if K and M and N else ""
        stats["ok" if statut.startswith("OK") else "verif"] += 1
        reste = (time.time() - debut) / (n + 1) * (len(a_traiter) - n - 1)
        print(f"    => {(VERT if statut.startswith('OK') else ROUGE)}{nouveau or '(pas de code)'}  [{statut}]{FIN}  "
              f"(reste ~{reste / 60:.1f} min)", flush=True)
        rapport.append([chemin.name, r, i["num"], i["nim"], ancien, nouveau, statut])

    # Collisions : doublons du code complet sur TOUT le fichier (+ couleur "collision" de la legende)
    groupes, teinte = {}, next(iter(roles["collision"]), None)
    for i in infos:
        ws.cell(i["ligne"], col["collision"]).value = None
        code = code_de(ligne_info(ws, i["ligne"], col))
        if code:
            groupes.setdefault(code, []).append(i["ligne"])
    nb_coll = 0
    for lignes in groupes.values():
        if len(lignes) > 1:
            nb_coll += len(lignes)
            for r in lignes:
                ws.cell(r, col["collision"]).value = f"OUI (lignes {', '.join(map(str, lignes[:8]))}{'...' if len(lignes) > 8 else ''})"
                if teinte:
                    ws.cell(r, col["collision"]).fill = PatternFill("solid", fgColor="FF" + teinte)
    sortie = DOSSIER_OUTPUT / f"{chemin.stem}_codifie_{horodatage}.xlsx"
    wb.save(sortie)
    print(f"\nResume : {len(a_traiter)} traitees | {stats['ok']} OK | {stats['verif']} a verifier | "
          f"{nb_coll} lignes en collision | {(time.time() - debut) / 60:.1f} min\nFichier : {sortie}\n")


def main():
    os.system("")  # active les couleurs ANSI sous Windows
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    DOSSIER_INPUT.mkdir(exist_ok=True)
    DOSSIER_OUTPUT.mkdir(exist_ok=True)
    fichiers = sorted(p for p in DOSSIER_INPUT.glob("*.xlsx") if not p.name.startswith("~$"))
    if not fichiers:
        sys.exit(f"Le dossier '{DOSSIER_INPUT.name}' ne contient aucun fichier .xlsx : depose ton fichier Excel dedans.")
    verifier_ollama()
    maintenant, rapport = datetime.now(), []
    chemin_cache = DOSSIER_OUTPUT / "cache_embeddings.json"
    try:
        cache = json.loads(chemin_cache.read_text())
    except (OSError, ValueError):
        cache = {}
    for f in fichiers:
        print(f"\n=== {f.name} ===")
        traiter_fichier(f, maintenant.strftime("%Y-%m-%d_%Hh%M"), rapport, cache)
        chemin_cache.write_text(json.dumps(cache))
    chemin_csv = DOSSIER_OUTPUT / f"rapport_{maintenant.strftime('%Y-%m-%d')}.csv"
    with open(chemin_csv, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["fichier", "ligne", "N°", "designation NIMRODA", "ancien code", "nouveau code", "statut"])
        w.writerows(rapport)
    print(f"Rapport : {chemin_csv}")


if __name__ == "__main__":
    main()
