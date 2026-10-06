# Codification automatique d'articles consommables

Codifie les lignes à traiter d'un fichier Excel (onglets `Legende`, `Codification`, `Referentiel`) avec une IA locale (Ollama, `qwen2.5:3b`), pensée pour un petit PC (i3, 8 Go de RAM).

## Mode d'emploi
1. Installe Python 3.10+ et [Ollama](https://ollama.com), puis lance `ollama pull qwen2.5:3b` (Ollama doit tourner).
2. Dépose ton fichier `.xlsx` dans le dossier `input/`.
3. Double-clique sur `lancer.bat` : le fichier d'origine n'est jamais modifié.
4. Récupère `output/<fichier>_codifie_<date>.xlsx` et `output/rapport_<date>.csv`.
5. Relis les lignes jaunes (statuts `A VERIFIER` en priorité), puis enlève toi-même le jaune.

## Comment ça marche
- **Légende** : le programme lit l'onglet `Legende` (couleur + signification) pour repérer les cellules « à modifier » et afficher le sens des couleurs de chaque ligne. Les colonnes sont retrouvées par leur en-tête.
- **Lignes à traiter** : colonne A en jaune, ou Code Famille / Sous-famille / Détail incomplets.
- **Referentiel dynamique** : relu à chaque lancement. Une nouvelle famille ou sous-famille (ex. « Téléviseurs ») est trouvée grâce à son libellé, sans toucher au code. Les plages `VLOOKUP` figées du fichier (ex. `$A$2:$B$7`) sont élargies dans la copie de sortie.
- **Sous-famille** : choisie par recherche de mots (sans IA) quand c'est net (`OK`). Sinon l'IA tranche entre 6 candidates seulement ; son choix n'est validé que s'il confirme celui de la recherche, sinon `A VERIFIER : sous-famille incertaine`.
- **Détail (colonne N)** : jamais écrit par l'IA, qui se trompait trop. Il est construit par gabarit : chaque ligne déjà validée proche « vote » en appliquant sa forme (`D<diamètre>-G<grain>`, `T<taille>`…) aux nombres de l'article (désignation fournisseur d'abord). Si tous les votes concordent : `OK` ; sinon `A VERIFIER : detail incertain`. Sans voisin : règles génériques, puis `STD`.
- **Autres statuts** : `A VERIFIER : collision` (code déjà pris), `code trop long`, `sous-famille introuvable`. Colonne Collision recalculée sur tout le fichier.

## Conseils pour un PC lent
Branche-le sur secteur, ferme Excel et le navigateur pendant le traitement. Dans `codifier.py`, `MODELE = "qwen2.5:1.5b"` est environ 2 fois plus rapide mais moins précis.
