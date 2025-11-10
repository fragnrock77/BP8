# Explorateur de données CSV & XLSX

## Résumé
Cette application web permet de charger des fichiers CSV ou Excel volumineux directement dans le navigateur, d'effectuer des recherches booléennes avancées sur les lignes et d'exporter les résultats filtrés. Elle propose également un mode de comparaison qui confronte un fichier de référence à un fichier à analyser et indique, pour chaque ligne, les mots-clés retrouvés. L'outil fonctionne entièrement côté client : aucune donnée n'est envoyée vers un serveur, ce qui garantit la confidentialité des informations analysées.

## Fonctionnalités principales
- Import de fichiers CSV, XLS ou XLSX jusqu'à 200 Mo via glisser-déposer ou sélection manuelle.
- Mode comparaison : import d'un fichier source de référence et d'un fichier à analyser, extraction automatique des mots-clés du fichier source et affichage, ligne par ligne, des correspondances trouvées.
- Lecture progressive des CSV grâce à PapaParse et prise en charge des classeurs Excel via SheetJS.
- Recherche multi-mots avec opérateurs `AND`, `OR`, `NOT`, guillemets, sensibilité à la casse et correspondance exacte.
- Pagination de 100 lignes, compteur de résultats et affichage d'un tableau réactif.
- Copie des résultats dans le presse-papiers, export CSV ou Excel.
- Exécution locale sans dépendance serveur : idéal pour les données confidentielles.

## Prise en main rapide
1. Ouvrez `index.html` dans un navigateur moderne.
2. Choisissez le mode "Analyse d'un fichier" ou "Comparaison de fichiers" puis importez le ou les fichiers en les déposant dans la zone prévue ou via les boutons.
3. Utilisez la zone de recherche pour filtrer les lignes (voir la section [Syntaxe de recherche](#syntaxe-de-recherche)).
4. Parcourez les pages de résultats, copiez ou exportez les lignes filtrées si nécessaire.

### Mode comparaison
1. Sélectionnez un **fichier de référence** contenant les mots-clés à repérer. Les valeurs non vides de toutes les colonnes sont dédupliquées et utilisées comme lexique.
2. Importez ensuite le **fichier à comparer** : chaque ligne reçoit une colonne « Mots-clés trouvés » listant les entrées du fichier de référence détectées, en respectant les options de sensibilité à la casse et de correspondance exacte.
3. Affinez éventuellement les résultats avec la recherche booléenne, puis copiez ou exportez le tableau enrichi.

## Syntaxe de recherche
- **Opérateurs booléens** : utilisez `AND`, `OR`, `NOT` (insensibles à la casse) pour combiner les mots-clés.
- **Groupement** : encadrez les sous-expressions avec des parenthèses, par exemple `(premium AND NOT résilié) OR VIP`.
- **Expression exacte** : entourez une phrase de guillemets `"..."` pour rechercher la séquence exacte.
- **Sensibilité à la casse** : activez l'option dédiée pour distinguer majuscules et minuscules.
- **Correspondance exacte** : cochez "Correspondance exacte" pour trouver uniquement des mots entiers.

## Architecture
| Fichier | Rôle |
| --- | --- |
| `index.html` | Structure HTML de la page, intégration des scripts tiers (PapaParse & SheetJS) et de `app.js`. |
| `styles.css` | Thème clair/sombre, styles des zones de dépôt, formulaires, tableau et pagination. |
| `app.js` | Logique applicative : import simple ou comparatif, parsing, extraction des mots-clés de référence, moteur de recherche booléen, pagination et exports. |
| `tests/run-tests.js` | Suite de tests Node.js validant le moteur de recherche et la génération CSV. |

## Installation & développement
```bash
npm install
```
Aucune dépendance n'est installée côté client : PapaParse et SheetJS sont chargés depuis un CDN.

### Lancer les tests
```bash
node tests/run-tests.js
```
Les tests vérifient la tokenisation de la requête, la conversion en notation postfixée, l'évaluation booléenne, la génération de CSV, la cohérence des caches de recherche ainsi que la normalisation des données importées et l'extraction des mots-clés de référence.

## Limitations actuelles
- Le traitement se fait en mémoire : les fichiers dépassant 200 Mo sont refusés pour préserver la réactivité du navigateur.
- L'application ne persiste pas l'état : un rafraîchissement de page réinitialise les données.

## Ressources complémentaires
- [Documentation PapaParse](https://www.papaparse.com/docs)
- [Documentation SheetJS](https://docs.sheetjs.com/)
