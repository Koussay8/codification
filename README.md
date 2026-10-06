# Codification automatique d'articles consommables

Codifie les lignes à traiter d'un fichier Excel (onglets `Legende`, `Codification`, `Referentiel`) avec une IA locale (Ollama), pensée pour un petit PC (i3, 8 Go de RAM).

## Mode d'emploi
1. Installe Python 3.10+ et [Ollama](https://ollama.com) (version 0.9 ou plus récente), puis lance `ollama pull qwen3:4b` et `ollama pull qwen3-embedding:0.6b` (Ollama doit tourner).
2. Dépose ton fichier `.xlsx` dans le dossier `input/`.
3. Double-clique sur `lancer.bat` : le fichier d'origine n'est jamais modifié.
4. Récupère `output/<fichier>_codifie_<date>.xlsx` et `output/rapport_<date>.csv`.
5. Relis les lignes jaunes (statuts `A VERIFIER` en priorité), puis enlève toi-même le jaune.

## Comment ça marche
- **Légende** : le programme lit l'onglet `Legende` (couleur + signification) pour repérer les cellules « à modifier » et afficher le sens des couleurs de chaque ligne. Les colonnes sont retrouvées par leur en-tête.
- **Lignes à traiter** : colonne A en jaune, ou Code Famille / Sous-famille / Détail incomplets.
- **Referentiel dynamique** : relu à chaque lancement. Une nouvelle famille ou sous-famille (ex. « Téléviseurs ») est trouvée grâce à son libellé, sans toucher au code. Les plages `VLOOKUP` figées du fichier (ex. `$A$2:$B$7`) sont élargies dans la copie de sortie.
- **Sous-famille** : un petit modèle d'embeddings (`qwen3-embedding:0.6b`) compare le *sens* de l'article au libellé de chaque sous-famille du Referentiel et aux lignes déjà validées (pas de recherche par mots). Si c'est net, la sous-famille est retenue (`OK`). Sinon le LLM (`qwen3:4b`) tranche entre 4 à 6 candidates avec les exemples les plus proches ; son choix n'est validé que s'il confirme l'un des deux signaux, sinon `A VERIFIER : sous-famille incertaine`. Une sous-famille ajoutée au Referentiel est reconnue grâce à son libellé, même sans exemple. Si le format du détail ne s'applique pas à l'article, la candidate suivante est essayée. Les vecteurs sont mis en cache (`output/cache_embeddings.json`).
- **Détail (colonne N)** : jamais écrit par l'IA, qui inventait. Des règles lisent dans les désignations (fournisseur d'abord) le diamètre, les cotes, le grain, la taille, le volume, puis appliquent le format de la sous-famille (`D<d>.<l>-G<grain>` pour les roues/manchons, `D<d>G<grain>` pour les disques, `T<taille>` pour les gants, `D<diamètre>` pour les forets, `STD`…). Le format de chaque sous-famille est appris de ton fichier : d'abord les lignes validées, sinon les codes déjà présents sur les lignes à traiter, sinon l'IA choisit par analogie (statut `A VERIFIER`).
- **Autres statuts** : `A VERIFIER : collision` (code déjà pris), `attribut introuvable`, `detail propre au produit` (codes comme `TMOY` ou `H540A` qu'aucune règle ne déduit), `code trop long`, `sous-famille introuvable`. Colonne Collision recalculée sur tout le fichier.

## Conseils pour un PC lent
Branche-le sur secteur, ferme Excel et le navigateur pendant le traitement. Le LLM n'est appelé que pour les lignes ambiguës (environ 40 %). Dans `codifier.py`, `MODELE = "qwen2.5:3b"` est un peu plus léger mais moins précis.
