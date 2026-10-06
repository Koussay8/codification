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
- **Détail (colonne N)** : jamais écrit par l'IA, qui inventait. Des règles lisent dans les désignations (fournisseur d'abord) le diamètre, les cotes, le grain, la taille, le volume, puis appliquent le format de la sous-famille (`D<d>.<l>-G<grain>` pour les roues/manchons, `D<d>G<grain>` pour les disques, `T<taille>` pour les gants, `D<diamètre>` pour les forets, `STD`…). Le format de chaque sous-famille est appris de ton fichier : d'abord les lignes validées, sinon les codes déjà présents sur les lignes à traiter, sinon l'IA choisit par analogie (statut `A VERIFIER`).
- **Autres statuts** : `A VERIFIER : collision` (code déjà pris), `attribut introuvable`, `detail propre au produit` (codes comme `TMOY` ou `H540A` qu'aucune règle ne déduit), `code trop long`, `sous-famille introuvable`. Colonne Collision recalculée sur tout le fichier.

## Conseils pour un PC lent
Branche-le sur secteur, ferme Excel et le navigateur pendant le traitement. Dans `codifier.py`, `MODELE = "qwen2.5:1.5b"` est environ 2 fois plus rapide mais moins précis.
