# CSV → Excel

Ouvre les fichiers CSV dans Excel par double-clic, avec les bons réglages détectés automatiquement, au lieu du séparateur et de l'encodage imposés par Windows.

## Avant / après

Avant : double-clic sur un CSV, Excel applique le séparateur et l'encodage de Windows.

![Excel ouvre mal le CSV](https://files.italic.fr/vuBHv-2026-11/EXCEL_HqSWgqGkPT.png)

Après : même fichier, ouvert via le script.

![CSV correctement importé par le script](https://files.italic.fr/gC3Ww-2026-11/EXCEL_yakeQJKI1i.png)

## Détection

- Encodage : UTF-8 (avec ou sans BOM), UTF-16, ANSI
- Séparateur : `;` `,` tabulation `|`, ligne `sep=x` respectée
- Séparateur décimal : `.` ou `,`
- Colonnes gardées en texte : zéros en tête, `+33…`, nombres de plus de 15 chiffres

Les données sont importées dans un nouveau classeur, sans lien vers le fichier source.

## Installation

Prérequis : Windows 10/11, Excel, Windows PowerShell 5.1 (inclus dans Windows).

1. Télécharger le dépôt et lancer `install.cmd` (pas de droits administrateur, installation pour l'utilisateur courant dans `%LOCALAPPDATA%\CsvExcel`).
2. Clic droit sur un `.csv` → **Ouvrir avec** → **Choisir une autre application** → **Excel (import CSV propre)** → **Toujours**.

Un menu « Ouvrir dans Excel (import propre) » est aussi ajouté au clic droit sur les CSV (sous Windows 11 : « Afficher d'autres options »).

![Menu clic droit « Ouvrir dans Excel (import propre) »](https://files.italic.fr/udAi5-2026-11/explorer_XexVtcWYn9.png)

## Désinstallation

Lancer `uninstall.cmd`.

## Limites

- Brève apparition possible d'une fenêtre PowerShell à l'ouverture.
- L'enregistrement propose `.xlsx` par défaut, pas le CSV d'origine.
- Excel ne gère pas plus de 1 048 576 lignes.

## Licence

[MIT](LICENSE)
