# scraper_cpf

Petit utilitaire qui lit un fichier Excel exporté de MonCompteFormation, ouvre chaque URL détail dans Chrome (Puppeteer) et copie automatiquement le bloc « Contenu » dans une colonne dédiée.

## Prérequis
- Node.js 18+
- `npm install`
- Un fichier `.xlsx` avec une feuille `Formations` contenant une colonne `URL détail`

## Utilisation rapide
1. Déposez votre fichier Excel dans le dossier `input/` (ou indiquez son chemin complet).
2. Lancez l’outil :
   ```bash
   npm run extract -- --file=input/mon_fichier.xlsx
   ```
   Sans option `--file`, le script prend automatiquement le fichier `.xlsx` le plus récent présent dans `input/`.
3. Le script lance Chrome (visible par défaut), visite chaque lien, récupère le bloc HTML « Contenu » (balises `p`, `ul`, `li`, etc.) et remplit la colonne correspondante.  
   Les lignes où la cellule « Contenu » est déjà remplie sont automatiquement ignorées.

## Options CLI
- `--file` / `-f` : chemin du fichier Excel à enrichir.
- `--help` : affiche l’aide.

## Configuration
Variables d’environnement utiles :

```
PUPPETEER_HEADLESS=false      # laisser à false pour voir le navigateur
NAVIGATION_TIMEOUT_MS=45000   # timeout d'une navigation (ms)
CONTENT_LOAD_DELAY_MS=2000    # délai après chargement avant extraction
BETWEEN_DELAY_MS=500          # pause entre deux fiches
```

Les logs sont écrits dans `logs/` (via Winston) ainsi qu’en console lorsque `NODE_ENV` ≠ production.

## Découpage d’un export volumineux

Lorsque le fichier `.xlsx` issu de MonCompteFormation devient trop lourd, utilisez l’utilitaire de découpe pour réduire sa taille sans perdre le lien entre les feuilles `Centres` et `Formations` :

```bash
npm run split -- --file=input/export_cpf.xlsx --max-centres=1000 --output-dir=output/split
```

- `--file` / `-f` : fichier source (par défaut le dernier `.xlsx` trouvé dans `input/`).
- `--max-centres` / `-c` : nombre maximal de centres par fichier (1000 par défaut).
- `--max-formations` / `-m` : limite additionnelle sur le nombre de formations.
- `--output-dir` / `-o` : dossier de sortie (défaut `./output`).
- `--prefix` / `-p` : préfixe des fichiers générés (`<nom_source>_part-XX.xlsx` par défaut).

Chaque fichier généré contient uniquement les centres sélectionnés et toutes les formations associées, plus un onglet `Résumé` avec les compteurs principaux pour cette partie.
