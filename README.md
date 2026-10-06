# Codification automatique d'articles consommables

Codifie les lignes à traiter d'un fichier Excel avec une IA locale (Ollama, `qwen2.5:3b`).

## Mode d'emploi
1. Installe Python 3.10+ et [Ollama](https://ollama.com), puis lance `ollama pull qwen2.5:3b` (Ollama doit tourner).
2. Dépose ton fichier `.xlsx` (onglets `Legende`, `Codification`, `Referentiel`) dans le dossier `input/`.
3. Double-clique sur `lancer.bat` : le fichier d'origine n'est jamais modifié.
4. Récupère `output/<fichier>_codifie_<date>.xlsx` et `output/rapport_<date>.csv`.
5. Relis les lignes jaunes (statut `A VERIFIER` en priorité), puis enlève toi-même le jaune.
