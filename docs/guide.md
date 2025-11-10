# Guide de l'application

## Vue d'ensemble
L'Explorateur de données CSV & XLSX est une application monopage qui s'exécute entièrement dans le navigateur. Elle vise à faciliter l'analyse ponctuelle de fichiers volumineux sans exposer les données à un serveur externe. Trois briques principales composent l'outil :

1. **Chargement & parsing** – prise en charge des formats CSV et Excel.
2. **Moteur de recherche** – filtrage booléen efficace basé sur des caches textuels.
3. **Restitution** – pagination, tableau dynamique, export et copie.

## Flux utilisateur
1. **Choix du mode & import** : `setMode` bascule entre l'analyse d'un seul fichier et la comparaison de deux fichiers.
   - En mode simple, `handleSingleFile` valide le fichier (`getFileExtension`), déclenche le parsing et remet à zéro l'état précédent.
   - En mode comparatif, `handleReferenceFiles` lit le fichier source de référence, `extractKeywords` en déduit un lexique unique, puis `handleComparisonFiles` prépare les lignes du fichier à analyser. Le glisser-déposer de deux fichiers est géré par `handleCompareDrop` (premier fichier = référence, second = comparaison).
2. **Parsing & normalisation** :
   - Pour un CSV, `parseCsv` exploite PapaParse en mode streaming (`step`) afin de supporter de gros volumes et d'alimenter la barre de progression.
   - Pour un classeur Excel (`.xlsx` ou `.xls`), `parseXlsx` s'appuie sur SheetJS pour convertir la première feuille en matrice.
   - `normalizeParsedData` convertit chaque cellule en chaîne de caractères, récupère les en-têtes et prépare les tableaux utilisés par l'application.
3. **Construction de l'état** : `updateComparisonDataset` ajoute la colonne « Mots-clés trouvés » en comparant chaque ligne du fichier cible au lexique extrait. Dans tous les cas, `applyDataset` hydrate `rawRows`, construit les caches (`buildCaches`) et affiche les sections de recherche/résultats.
4. **Recherche** : l'utilisateur lance `performSearch` qui tokenise la requête (`tokenizeQuery`), la convertit en notation postfixée (`toPostfix`) puis évalue l'expression (`evaluateQuery`). Les index correspondants alimentent `filteredRows`.
5. **Affichage** : `renderPage` construit le tableau HTML via `renderTable`, affiche les statistiques et gère la pagination (`prevPageBtn`, `nextPageBtn`).
6. **Export** : `copyToClipboard`, `exportCsv` et `exportXlsx` exploitent `convertRowsToCsv` ou SheetJS pour diffuser les résultats.

## Modules clés (`app.js`)
| Fonction | Description |
| --- | --- |
| `setMode` | Affiche les panneaux d'import adéquats, réinitialise l'état lors du changement de mode. |
| `handleSingleFile` | Valide le fichier, le parse puis charge les résultats dans le tableau. |
| `handleReferenceFiles` / `handleComparisonFiles` | Gèrent respectivement le fichier source de référence et le fichier à comparer, puis mettent à jour l'état. |
| `handleCompareDrop` | Dispatch des fichiers déposés pour alimenter la comparaison. |
| `parseCsv` | Lit les CSV en streaming, gère les erreurs et alimente les barres de progression. |
| `parseXlsx` | Convertit la première feuille Excel en tableau de lignes. |
| `normalizeParsedData` | Nettoie les valeurs, détecte les en-têtes et renvoie des tableaux prêts à l'emploi. |
| `extractKeywords` | Produit un ensemble de mots-clés uniques à partir du fichier de référence. |
| `updateComparisonDataset` | Construit les lignes enrichies avec la colonne « Mots-clés trouvés ». |
| `buildCaches` | Construit des versions concaténées des lignes pour accélérer `matchRow`. |
| `tokenizeQuery` | Découpe la requête utilisateur en opérandes/opérateurs, en tenant compte des guillemets. |
| `toPostfix` | Transforme l'expression infixée en notation postfixée (algorithme de shunting-yard). |
| `evaluateQuery` | Évalue la requête postfixée pour chaque ligne en appliquant `matchRow`. |
| `renderTable` | Génére le `<table>` et assure la robustesse face aux colonnes manquantes. |
| `renderPage` | Découpe `filteredRows`, met à jour la pagination et les statistiques. |
| `convertRowsToCsv` | Sérialise les lignes filtrées en CSV (échappement des guillemets et virgules). |

## Gestion des erreurs
- Taille maximale contrôlée (`MAX_FILE_SIZE`).
- Vérification d'extension autorisée.
- Messages d'erreur utilisateur via `showError`.
- Gestion des erreurs de parsing (abort PapaParse, exceptions SheetJS).
- Mise à jour de la barre de progression pour distinguer le chargement du fichier de référence, du fichier comparé et de la comparaison finale.

## Accessibilité & UX
- Boutons et entrées conformes aux standards (focus visible, rôle `alert` pour les erreurs).
- Barre de progression et libellés textuels pour informer l'utilisateur.
- Pagination et indicateurs de résultats accessibles (`aria-live` sur la table).

## Tests automatisés
Le fichier `tests/run-tests.js` s'exécute avec Node.js et vérifie :
- La tokenisation correcte des requêtes complexes.
- La gestion des parenthèses et la priorité des opérateurs.
- L'évaluation des expressions booléennes avec options de casse/correspondance exacte.
- La génération correcte du CSV, la synchronisation des caches de recherche, la normalisation des données importées et l'extraction de mots-clés uniques.

## Personnalisation
- **Taille des pages** : modifier `PAGE_SIZE` dans `app.js`.
- **Limite de taille des fichiers** : ajuster `MAX_FILE_SIZE`.
- **Thème** : les couleurs principales sont définies dans `styles.css` via les variables CSS.

## Déploiement
Étant une application statique, un simple hébergement de fichiers (GitHub Pages, Netlify, S3, etc.) suffit. Veillez à autoriser le chargement des scripts CDN utilisés pour PapaParse et SheetJS.
