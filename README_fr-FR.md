[English](README.md) | [简体中文](README_zh-CN.md) | [繁體中文](README_zh-TW.md) | [Deutsch](README_de-DE.md) | [Français](README_fr-FR.md) | [Русский](README_ru-RU.md) | [Português](README_pt-BR.md) | [日本語](README_ja-JP.md) | [한국어](README_ko-KR.md) | [Español](README_es-ES.md) | [Tiếng Việt](README_vi-VN.md)

# DocWen

Un logiciel de conversion de format de documents et de graphiques - Prend en charge la conversion bidirectionnelle Word/Markdown/Excel. Fonctionne en local, assurant la sécurité et la fiabilité des données.

## 📖 Contexte du projet

Ce logiciel a été conçu à l'origine pour le travail quotidien du service d'impression afin de résoudre les problèmes suivants :
- Les formats de documents envoyés par divers départements sont chaotiques et doivent être organisés dans des formats standardisés.
- Il existe de nombreux types de documents, chacun avec des exigences de format fixes différentes.
- Doit fonctionner hors ligne, s'adaptant aux environnements intranet et aux équipements anciens.

**Philosophie de conception** : Ce logiciel se positionne comme un outil léger et simple. Bien qu'il ne puisse pas être comparé à des outils professionnels comme LaTeX ou Pandoc en termes de professionnalisme et d'exhaustivité fonctionnelle, il excelle par son coût d'apprentissage nul et sa facilité d'utilisation immédiate, ce qui le rend adapté aux scénarios de bureau quotidiens où les exigences de format ne sont pas extrêmement strictes.

## ✨ Fonctionnalités principales

- **📄 Conversion de format de document** - Conversion bidirectionnelle Word ↔ Markdown. Prend en charge la conversion de formules mathématiques et la conversion bidirectionnelle des séparateurs (les trois types de lignes de séparation de Markdown vs sauts de page, sauts de section et lignes horizontales de Word). Prend en charge les formats tels que DOCX/DOC/WPS/RTF/ODT.
- **📊 Conversion de format de feuille de calcul** - Conversion bidirectionnelle Excel ↔ Markdown. Prend en charge les formats XLSX/XLS/ET/ODS/CSV. Inclut des outils de résumé de tableau.
- **📑 PDF et fichiers de mise en page** - Conversion PDF/XPS/OFD vers Markdown ou DOCX. Prend en charge la fusion, la division de PDF et d'autres opérations.
- **🖼️ Traitement d'image** - Prend en charge la conversion bidirectionnelle et la compression des formats JPEG/PNG/GIF/BMP/TIFF/WebP/HEIC.
- **🔍 Reconnaissance de texte OCR** - RapidOCR intégré pour extraire du texte à partir d'images et de PDF.
- **✏️ Correction de texte** - Vérifie les fautes de frappe, la ponctuation, les symboles et les mots sensibles en fonction de dictionnaires personnalisés. Les règles peuvent être modifiées dans l'interface des paramètres.
- **📝 Système de modèles** - Mécanisme de modèle flexible prenant en charge les formats de documents et de rapports personnalisés.
- **💻 Fonctionnement en double mode** - Interface utilisateur graphique (GUI) + Interface en ligne de commande (CLI).
- **🔒 Fonctionnement local** - Fonctionne hors ligne, assurant la sécurité des données avec des mécanismes d'isolation réseau intégrés.
- **🔗 Fonctionnement à instance unique** - Gère automatiquement les instances de programme et prend en charge l'intégration avec le plugin Obsidian associé.

## Journal des modifications

### v0.6.0 (2025-01-20)

- Prise en charge complète de l'internationalisation (GUI et CLI supportent 11 langues).
- Remplacement de PaddleOCR par RapidOCR pour une meilleure compatibilité.
- Ajout de modèles Word/Excel multilingues.
- Détection et injection automatiques des styles de modèle.
- Autres optimisations et corrections.

### v0.5.1 (2025-01-01)

- Ajout de la conversion bidirectionnelle des formules mathématiques (Word OMML ↔ Markdown LaTeX).
- Ajout de la conversion bidirectionnelle des notes de bas de page/notes de fin.
- Ajout de styles de caractères et de paragraphes pour le code, les citations, etc.
- Amélioration du traitement des listes (imbrication à plusieurs niveaux, numérotation automatique).
- Amélioration des fonctions de tableau (détection/injection de style, tableaux à trois lignes, etc.).
- Optimisation du nettoyage et de l'ajout des numéros de sous-titres.
- Amélioration de l'interaction de l'interface et de la liaison des paramètres.

### v0.4.1 (2025-12-05)

- Refonte de la CLI pour améliorer l'expérience utilisateur.
- Ajout de la prise en charge de plus de types de documents.
- Implémentation de plus d'options configurables.

## 🚀 Démarrage rapide

### Lancer le programme

Double-cliquez sur `DocWen.exe` pour démarrer l'interface graphique.

### Guide de démarrage rapide

1.  **Préparer un fichier Markdown** :

    ```markdown
    ---
    Titre: Document de test
    ---
    
    ## Titre de test
    
    Ceci est le contenu du corps du test.
    ```

2.  **Conversion par glisser-déposer** :
    - Lancez le programme.
    - Faites glisser le fichier `.md` dans la fenêtre.
    - Sélectionnez un modèle.
    - Cliquez sur "Convertir en DOCX".

3.  **Obtenir les résultats** :
    - Un document Word standardisé sera généré dans le même répertoire.

**Conseil** : Vous pouvez utiliser les fichiers exemples dans le répertoire `samples/` pour essayer rapidement les fonctionnalités du logiciel.

## 📝 Conventions de syntaxe Markdown

### Mappage des niveaux d'en-tête

Pour faciliter la mémorisation pour les collègues sans connaissances de base, les en-têtes Markdown dans ce logiciel correspondent **un à un** aux en-têtes Word :
- Le titre du document (title) et le sous-titre (subtitle) sont placés dans les métadonnées YAML.
- Markdown `# En-tête 1` correspond à Word "Titre 1".
- Markdown `## En-tête 2` correspond à Word "Titre 2".
- Et ainsi de suite, prenant en charge jusqu'à 9 niveaux d'en-tête.

**Astuce** : Si vous préférez utiliser l'en-tête de premier niveau de Markdown (`#`) comme titre du document, en commençant par les en-têtes de deuxième niveau (`##`) pour les sous-titres du corps, vous pouvez styliser « Titre 1 » dans le modèle Word pour qu'il ressemble à un titre de document (par exemple, centré, gras, plus grande taille de police), et sélectionner un schéma de numérotation qui ignore la numérotation des en-têtes de premier niveau dans les paramètres. Ainsi, vos en-têtes de premier niveau apparaîtront comme des titres de document.

### Sauts de ligne et paragraphes

**Règle de base** : Chaque ligne non vide est traitée par défaut comme un paragraphe séparé.

**Paragraphes mixtes** : Lorsqu'un sous-titre doit être mélangé avec le corps du texte dans le même paragraphe, les conditions suivantes doivent être remplies :
1.  Le sous-titre se termine par un signe de ponctuation de fin (prend en charge la ponctuation multilingue, y compris les points, les points d'interrogation, les points d'exclamation et autres signes de ponctuation de fin courants).
2.  Le corps du texte est situé sur la **ligne immédiatement suivante** du sous-titre.
3.  La ligne du corps du texte ne peut pas être un élément Markdown spécial (comme les en-têtes, blocs de code, tableaux, listes, citations, blocs de formule, séparateurs, etc.).

**Exemple** :
```markdown
## I. Exigences de travail.
Cette réunion exige que toutes les unités mettent sérieusement en œuvre...
```
Les deux lignes ci-dessus seront fusionnées dans le même paragraphe, où "I. Exigences de travail." conserve le format de sous-titre, et "Cette réunion..." conserve le format du corps du texte.

**Remarque** :
- Il ne peut pas y avoir de ligne vide entre le sous-titre et le corps du texte ; sinon, ils seront reconnus comme des paragraphes séparés.
- Si le sous-titre ne se termine pas par un signe de ponctuation et qu'il n'y a pas de ligne vide avec le corps du texte, le corps du texte sera fusionné dans la ligne d'en-tête avec un formatage ajusté.

### Conversion bidirectionnelle des séparateurs

Prend en charge la conversion bidirectionnelle entre les séparateurs Markdown et les sauts de page/sauts de section/lignes horizontales Word :

-   **DOCX → MD** : Les sauts de page, sauts de section et lignes horizontales Word sont automatiquement convertis en séparateurs Markdown.
-   **MD → DOCX** : Markdown `---`, `***`, `___` sont automatiquement convertis en éléments Word correspondants.
-   **Configurable** : Les relations de mappage spécifiques peuvent être personnalisées dans l'interface des paramètres.

## 📖 Guide d'utilisation détaillé

### Word vers Markdown

1.  Faites glisser le fichier `.docx` dans la fenêtre du programme.
2.  Le programme analyse automatiquement la structure du document.
3.  Génère un fichier `.md` contenant des métadonnées YAML.

**Formats pris en charge** :
-   `.docx` - Document Word standard.
-   `.doc` - Automatiquement converti en DOCX pour le traitement.
-   `.wps` - Document WPS automatiquement converti.

**Options d'exportation** :

| Option | Description |
| :--- | :--- |
| **Extraire les images** | Si coché, les images du document sont extraites dans le dossier de sortie et les liens d'image sont insérés dans le fichier MD. |
| **OCR d'image** | Si coché, effectue l'OCR sur les images et crée un fichier image .md (contenant le texte reconnu). |
| **Nettoyer numéros sous-titres** | Si coché, supprime les numéros avant les sous-titres (ex: "一、", "（一）", "1.", etc.) et les convertit en texte de titre pur. |
| **Ajouter numéros sous-titres** | Si coché, ajoute automatiquement des numéros en fonction des niveaux d'en-tête (le schéma de numérotation peut être configuré dans les paramètres). |

### Markdown vers Word

1.  Préparez un fichier `.md` avec un en-tête YAML.
2.  Faites-le glisser dans la fenêtre du programme et sélectionnez le modèle Word correspondant.
3.  Le programme remplit automatiquement le modèle et génère le document.

**Options de conversion** :

| Option | Description |
| :--- | :--- |
| **Nettoyer numéros sous-titres** | Si coché, supprime les numéros avant les sous-titres. |
| **Ajouter numéros sous-titres** | Si coché, ajoute automatiquement des numéros en fonction des niveaux d'en-tête. |

**Remarque** : S'il y a des paragraphes où les sous-titres et le corps du texte sont mélangés dans le document, des sauts de ligne stricts doivent être maintenus dans le fichier MD (voir "Sauts de ligne et paragraphes" ci-dessus).

### Traitement automatique des styles de modèle

Le convertisseur détecte et traite automatiquement les styles de modèle lors de la conversion Markdown → DOCX :

#### Classification des styles

**Style de paragraphe** : Appliqué à l'ensemble du paragraphe.

| Style | Comportement de détection | Injection si manquant | Source |
| :--- | :--- | :--- | :--- |
| En-tête (1~9) | Détecte le style de paragraphe | Styles d'en-tête de modèle | Word Intégré |
| Bloc de code | Détecte le style de paragraphe | Police Consolas + Fond gris | Défini par le logiciel |
| Citation (1~9) | Détecte le style de paragraphe | Fond gris + Bordure gauche | Défini par le logiciel |
| Bloc de formule | Détecte le style de paragraphe | Style spécifique formule | Défini par le logiciel |
| Séparateur (1~3) | Détecte le style de paragraphe | Style de paragraphe bordure inférieure | Défini par le logiciel |

**Style de caractère** : Appliqué au texte sélectionné.

| Style | Comportement de détection | Injection si manquant | Source |
| :--- | :--- | :--- | :--- |
| Code en ligne | Détecte le style de caractère | Police Consolas + Ombrage gris | Défini par le logiciel |
| Formule en ligne | Détecte le style de caractère | Style spécifique formule | Défini par le logiciel |

**Style de tableau** : Appliqué à l'ensemble du tableau.

| Style | Comportement de détection | Injection si manquant | Source |
| :--- | :--- | :--- | :--- |
| Tableau à trois lignes | Priorité config utilisateur | Définition style tableau à trois lignes | Défini par le logiciel |
| Tableau grille | Priorité config utilisateur | Définition style tableau grille | Défini par le logiciel |

**Définition de numérotation** : Utilisé pour les formats de liste.

| Type | Comportement de détection | Traitement si manquant |
| :--- | :--- | :--- |
| Numérotation de liste | Analyse les définitions de liste ordonnée/non ordonnée existantes dans le modèle | Utilise le préréglage décimal/puce |

#### Internationalisation des noms de style

-   **Styles intégrés Word** (heading 1~9) :
    -   Les noms de style utilisent les noms anglais standard de Word (ex: `heading 1`).
    -   Word affiche automatiquement les noms localisés en fonction de la langue du système (ex: "Titre 1" sur les systèmes français).
-   **Styles définis par le logiciel** (Bloc de code, Citation, Formule, Séparateur, Tableau, etc.) :
    -   Injecte les noms de style de langue correspondants en fonction du paramètre de langue de l'interface du logiciel.
    -   Interface chinoise : Injecte "代码块", "引用 1", "三线表", etc.
    -   Interface anglaise : Injecte "Code Block", "Quote 1", "Three Line Table", etc.

**Suggestion** : Après avoir personnalisé les styles dans le modèle, le convertisseur utilisera automatiquement vos styles ; s'ils ne sont pas présents dans le modèle, il utilisera les styles prédéfinis intégrés.

### Traitement des fichiers de feuille de calcul

1.  **Excel/CSV vers Markdown** : Faites glisser des fichiers `.xlsx` ou `.csv` pour convertir automatiquement en tableaux Markdown.
2.  **Markdown vers Excel** : Préparez un fichier MD et sélectionnez un modèle Excel pour la conversion.

**Formats pris en charge** :
-   `.xlsx` - Document Excel standard.
-   `.xls` - Automatiquement converti en XLSX pour le traitement.
-   `.et` - Feuille de calcul WPS automatiquement convertie.
-   `.csv` - Tableau texte CSV.

### Fonction de correction de texte

Le programme fournit quatre règles de correction personnalisables :

1.  **Vérification de l'appariement de la ponctuation** - Détecte si la ponctuation par paire comme les parenthèses et les guillemets correspond.
2.  **Correction de symboles** - Détecte l'utilisation mixte de la ponctuation chinoise et anglaise.
3.  **Vérification des fautes de frappe** - Vérifie les fautes de frappe courantes en fonction d'un dictionnaire personnalisé.
4.  **Détection de mots sensibles** - Détecte les mots sensibles en fonction d'un dictionnaire personnalisé.

**Dictionnaires personnalisés** : Modifiez visuellement les dictionnaires de fautes de frappe et de mots sensibles dans l'interface "Paramètres".

**Utilisation** :
1.  Faites glisser le document Word à vérifier dans le programme.
2.  Cochez les règles de correction requises.
3.  Cliquez sur le bouton "Correction de texte".
4.  Les résultats de la correction sont affichés sous forme de commentaires dans le document.

## 🛠️ Système de modèles

### Utilisation de modèles existants

Le programme est livré avec divers modèles, y compris des versions multilingues. Vous pouvez les sélectionner et les utiliser selon vos besoins. Les fichiers modèles sont situés dans le répertoire `templates/`.

### Modèles personnalisés

1.  Créez un fichier modèle à l'aide de Word ou WPS.
2.  Référez-vous aux modèles existants et insérez des espaces réservés comme `{{Title}}`, etc., où le remplissage est nécessaire.
3.  Dans le modèle, les styles intégrés Titre 1 ~ Titre 5 doivent être modifiés manuellement.
4.  Enregistrez le modèle dans le répertoire `templates/`.
5.  Redémarrez le programme, et le nouveau modèle sera automatiquement chargé.

Vous pouvez également copier un modèle existant, le modifier et le renommer.

### Utilisation des espaces réservés

#### Espaces réservés de modèle Word

**Espaces réservés de champ YAML** : Utilisez le format `{{NomChamp}}` dans le modèle, qui sera remplacé par la valeur correspondante dans l'en-tête YAML du fichier Markdown lors de la conversion.

| Espace réservé | Description |
| :--- | :--- |
| `{{Titre}}` | Titre du document (Règles de récupération voir ci-dessous)  |
| `{{Corps}}` | Position d'insertion du contenu du corps Markdown |
| Autres | Prend en charge tout champ personnalisé |

**Priorité de récupération du titre** :

| Priorité | Source | Description |
| :--- | :--- | :--- |
| 1 | Champ YAML `Title` | Priorité la plus élevée |
| 2 | Champ YAML `aliases` | Prend le premier élément de la liste, ou la valeur de chaîne |
| 3 | Nom de fichier | Nom de fichier sans extension `.md` |

**Support multilingue** : Les espaces réservés titre et corps supportent plusieurs langues, ex: titre peut être `{{Titre}}`, `{{title}}`, `{{标题}}`, etc., corps peut être `{{Corps}}`, `{{body}}`, `{{正文}}`, etc.

#### Espaces réservés de modèle Excel

Les modèles Excel prennent en charge trois types d'espaces réservés :

**1. Espace réservé de champ YAML** `{{NomChamp}}`

Utilisé pour remplir une valeur unique à partir de l'en-tête YAML du fichier Markdown :

```markdown
---
ReportName: Statistiques de ventes annuelles 2024
Unit: Service des ventes
---
```

`{{ReportName}}`, `{{Unit}}` dans le modèle seront remplacés par les valeurs correspondantes. Le champ titre suit également les règles de priorité.

**2. Espace réservé de remplissage de colonne** `{{↓NomChamp}}`

Extrait les données du tableau Markdown et remplit **vers le bas** ligne par ligne à partir de la position de l'espace réservé :

```markdown
| NomProduit | Quantité |
|:--- |:--- |
| Produit A | 100 |
| Produit B | 200 |
```

`{{↓NomProduit}}` dans le modèle Excel sera remplacé par "Produit A", et la ligne suivante sera remplie avec "Produit B".

**3. Espace réservé de remplissage de ligne** `{{→NomChamp}}`

Extrait les données du tableau Markdown et remplit **vers la droite** colonne par colonne à partir de la position de l'espace réservé :

```markdown
| Mois |
|:--- |
| Jan |
| Fév |
| Mar |
```

`{{→Mois}}` dans le modèle Excel sera rempli séquentiellement avec "Jan", "Fév", "Mar" vers la droite.

**Traitement des cellules fusionnées** : Le programme ignore automatiquement les cellules non premières des cellules fusionnées pour assurer un remplissage correct des données.

**Fusion de données multi-tableaux** : S'il y a plusieurs tableaux dans Markdown utilisant le même nom d'en-tête, les données seront fusionnées dans l'ordre et remplies séquentiellement.

## 🖥️ Utilisation de l'interface graphique

La plupart des utilisateurs utilisent ce logiciel via l'interface graphique. Voici le guide d'utilisation détaillé.

### Aperçu de l'interface

Le programme utilise une **mise en page adaptative à trois colonnes** :

| Zone | Description | Moment d'affichage |
| :--- | :--- | :--- |
| **Colonne centrale (Zone principale)** | Zone de glisser-déposer de fichier, panneau d'opération, barre d'état | Toujours affiché |
| **Colonne de droite** | Sélecteur de modèle / Panneau de conversion de format | S'étend automatiquement après la sélection d'un fichier |
| **Colonne de gauche** | Liste de fichiers par lots (groupés par type) | Affiché lors du passage en mode par lots |

### Flux d'opération de base

1.  **Lancer le programme** : Double-cliquez sur `DocWen.exe`.
2.  **Importer le fichier** :
    -   Méthode 1 : Faites glisser et déposez les fichiers directement dans la fenêtre.
    -   Méthode 2 : Cliquez sur le bouton "Ajouter" dans la zone de glisser-déposer pour sélectionner des fichiers.
3.  **Sélectionner le modèle** (si conversion nécessaire) : Le panneau de modèle droit s'étend automatiquement ; sélectionnez un modèle approprié.
4.  **Configurer les options** : Cochez les options de conversion/exportation requises dans le panneau d'opération.
5.  **Exécuter l'opération** : Cliquez sur le bouton de fonction correspondant (ex : "Export MD", "Convertir en DOCX", etc.).
6.  **Voir le résultat** : La barre d'état affiche la progression et les résultats ; cliquez sur l'icône 📍 pour localiser le fichier de sortie.

### Mode fichier unique vs Mode par lots

Le programme prend en charge deux modes de traitement, commutables via le bouton bascule dans la zone de glisser-déposer de fichier :

**Mode fichier unique** (Par défaut) :
-   Traite un fichier à la fois.
-   Interface simple, adaptée à une utilisation quotidienne.

**Mode par lots** :
-   Importe plusieurs fichiers simultanément.
-   La colonne de gauche affiche la liste des fichiers catégorisés (groupés par document/tableau/image, etc.).
-   Prend en charge l'ajout, la suppression et le tri par lots.
-   Cliquer sur un fichier dans la liste change la cible de l'opération actuelle.

### Fonctions du panneau d'opération

Le panneau d'opération ajuste automatiquement les options disponibles en fonction du type de fichier :

| Type de fichier | Opérations disponibles |
| :--- | :--- |
| Document Word | Export MD, Convertir PDF, Correction texte, OCR |
| Markdown | Convertir DOCX, Convertir PDF |
| Tableau Excel | Export MD, Convertir PDF, Résumé tableau |
| Fichier PDF | Export MD, Fusionner, Diviser, OCR |
| Fichier Image | Conversion format, Compression, OCR |

### Interface des paramètres

Cliquez sur le bouton ⚙️ dans le coin inférieur droit de la fenêtre pour ouvrir les paramètres :

-   **Général** : Thème d'interface, langue, opacité fenêtre.
-   **Conversion** : Valeurs par défaut pour diverses options de conversion.
-   **Sortie** : Répertoire de sortie par défaut, règles de nommage de fichier.
-   **Correction** : Modifiez les dictionnaires de fautes de frappe et de mots sensibles.
-   **Style** : Configurations de style de bloc de code, citation, tableau.

### Raccourcis

-   **Faire glisser fichier externe** : Faites glisser directement dans la fenêtre pour importer.
-   **Double-clic résultat barre d'état** : Ouvrir rapidement le répertoire du fichier de sortie.
-   **Clic droit élément modèle** : Ouvrir l'emplacement du fichier modèle.

---

## 🔧 Utilisation en ligne de commande

En plus de l'interface graphique, le programme fournit une interface en ligne de commande (CLI), adaptée aux scripts d'automatisation et aux scénarios de traitement par lots.

### Modes d'exécution

-   **Mode interactif** : Affiche un guide de menu après avoir passé un fichier, similaire à l'opération GUI.
-   **Mode Headless** : Exécutez directement en ajoutant le paramètre `--action`, adapté à l'appel de script.

### Exemples courants

```bash
# Mode interactif
DocWen.exe document.docx

# Exporter Word en Markdown (Extraire images + OCR)
DocWen.exe report.docx --action export_md --extract-img --ocr

# Markdown vers Word (Spécifier modèle)
DocWen.exe document.md --action convert --target docx --template "Nom du modèle"

# Conversion par lots (Sauter confirmation, continuer sur erreur)
DocWen.exe *.docx --action export_md --batch --yes --continue-on-error

# Correction de document
DocWen.exe document.docx --action validate --check-typo --check-punct

# Fusion/Division PDF
DocWen.exe *.pdf --action merge_pdfs
DocWen.exe report.pdf --action split_pdf --pages "1-3,5,7-10"
```

### Arguments principaux

| Argument | Description |
| :--- | :--- |
| `--action` | Type d'opération : `export_md`, `convert`, `validate`, `merge_pdfs`, `split_pdf` |
| `--target` | Format cible : `pdf`, `docx`, `xlsx`, `md` |
| `--template` | Nom du modèle (ex : `Nom du modèle`) |
| `--extract-img` | Extraire les images lors de l'exportation |
| `--ocr` | Activer la reconnaissance OCR |
| `--batch` | Mode de traitement par lots |
| `--yes` / `-y` | Sauter les invites de confirmation |
| `--continue-on-error` | Continuer le traitement de l'élément suivant en cas d'erreur |
| `--json` | Sortir le résultat au format JSON |
| `--quiet` / `-q` | Mode silencieux, réduire la sortie |

## 🔌 Plugin Obsidian

Le projet inclut un plugin Obsidian correspondant pour travailler en tandem avec le convertisseur :

### Fonctionnalités principales

-   **🚀 Lancement en un clic** - Icône de la barre latérale pour lancer rapidement le convertisseur.
-   **📂 Transfert automatique** - Passe automatiquement le chemin du fichier actuellement ouvert.
-   **🔄 Gestion d'instance unique** - Envoie automatiquement le fichier si le programme est déjà en cours d'exécution, pas besoin de redémarrer.
-   **💪 Récupération après plantage** - Détecte automatiquement l'état du processus et nettoie automatiquement les fichiers résiduels.

### Principe de fonctionnement

Le plugin interagit avec le convertisseur via IPC basé sur le système de fichiers :

1.  **Premier clic** → Lancer le convertisseur et passer le fichier actuel.
2.  **Cliquer à nouveau (Avec fichier)** → Remplacer par le nouveau fichier (Mode fichier unique).
3.  **Cliquer à nouveau (Sans fichier)** → Activer la fenêtre du convertisseur.

### Installation

Le plugin a été publié dans un référentiel séparé. Veuillez visiter [docwen-obsidian](https://github.com/ZHYX91/docwen-obsidian) pour les instructions d'installation et la dernière version.

## ❓ FAQ

### Que faire si la conversion échoue ?

-   Vérifiez si le fichier est occupé par un autre programme.
-   Confirmez que le format du fichier est correct.
-   Vérifiez les journaux d'erreurs dans le répertoire `logs/`.

### Le modèle ne s'affiche pas ?

-   Confirmez que les fichiers modèles sont dans le répertoire `templates/`.
-   Vérifiez si le fichier modèle est corrompu.
-   Redémarrez le programme pour recharger les modèles.

### La fonction de correction ne fonctionne pas ?

-   Confirmez que le document est au format .docx.
-   Vérifiez si le document contient du texte modifiable.
-   Confirmez que les règles de correction sont activées dans les paramètres.

### Format de sortie non conforme aux attentes ?

-   Le programme génère des documents basés sur les styles de modèle. Pour ajuster le format de sortie, modifiez les définitions de style directement dans le fichier modèle.
-   Les fichiers modèles sont situés dans le répertoire `templates/`.
-   Après modification des styles de modèle, tous les documents convertis avec ce modèle appliqueront les nouveaux styles.

### Les cellules de formule sont vides après la conversion Excel vers Markdown ?

C'est un comportement attendu. Le programme lit les **valeurs mises en cache** des cellules plutôt que les formules elles-mêmes.

**Raison technique** :
-   Dans les fichiers Excel, les cellules de formule stockent à la fois la formule et le dernier résultat calculé (valeur mise en cache).
-   Le programme utilise le mode `data_only=True`, qui ne récupère que les valeurs mises en cache.
-   Si le fichier n'a jamais été ouvert dans Excel (par exemple, généré par un programme), ou a été édité mais pas ré-enregistré, la valeur mise en cache sera vide.

**Solution** :
1.  Ouvrez le fichier dans Excel.
2.  Attendez que le calcul des formules soit terminé.
3.  Enregistrez le fichier.
4.  Convertissez à nouveau.

## 🔒 Fonctionnalités de sécurité

-   **Fonctionnement local** : Tout le traitement est effectué localement, aucune dépendance réseau.
-   **Isolation réseau** : Le mécanisme d'isolation réseau intégré empêche les fuites de données.
-   **Pas de téléchargement de données** : Les fichiers utilisateur ne sont jamais téléchargés sur aucun serveur.

## 📜 Licence

Ce projet est sous licence **GNU Affero General Public License v3.0 (AGPL-3.0)**.

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

-   Ce projet utilise PyMuPDF (sous licence AGPL-3.0), donc l'ensemble du projet est également sous licence AGPL-3.0.
-   Vous êtes libre d'utiliser, de modifier et de distribuer ce logiciel.
-   Si vous modifiez ce logiciel et fournissez des services sur un réseau, vous devez fournir le code source modifié aux utilisateurs.
-   Pour des informations détaillées sur la licence, veuillez consulter le fichier [LICENSE](LICENSE).

### Contact

-   **GitHub** : https://github.com/ZHYX91/docwen
-   **Contacter l'auteur** : zhengyx91@hotmail.com

---

**Auteur** : ZhengYX
