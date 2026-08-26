# DocWen

<p align="center">
  <img src="https://raw.githubusercontent.com/ZHYX91/docwen/main/assets/icon.svg" alt="DocWen logo" width="120">
</p>

[English](https://github.com/ZHYX91/docwen/blob/main/README.md) · [简体中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-CN.md) · [繁體中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-TW.md) · [Deutsch](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.de-DE.md) · [Français](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.fr-FR.md) · [Español](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.es-ES.md) · [Português](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.pt-BR.md) · [Русский](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ru-RU.md) · [日本語](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ja-JP.md) · [한국어](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ko-KR.md) · [Tiếng Việt](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.vi-VN.md)

Un logiciel de conversion de format de documents et de graphiques - Prend en charge la conversion bidirectionnelle Word/Markdown/Excel. Fonctionne en local, assurant la sécurité et la fiabilité des données.

## 📖 Contexte du projet

Ce logiciel a été conçu à l'origine pour le travail quotidien du service d'impression afin de résoudre les problèmes suivants :
- Les formats de documents envoyés par divers départements sont chaotiques et doivent être organisés dans des formats standardisés.
- Il existe de nombreux types de documents, chacun avec des exigences de format fixes différentes.
- Doit fonctionner hors ligne, s'adaptant aux environnements intranet et aux équipements anciens.

**Philosophie de conception** : Ce logiciel se positionne comme un outil léger et simple. Bien qu'il ne puisse pas être comparé à des outils professionnels comme LaTeX ou Pandoc en termes de professionnalisme et d'exhaustivité fonctionnelle, il excelle par son coût d'apprentissage nul et sa facilité d'utilisation immédiate, ce qui le rend adapté aux scénarios de bureau quotidiens où les exigences de format ne sont pas extrêmement strictes.

## ✨ Fonctionnalités principales

- **📄 Conversion de format de document** - Conversion bidirectionnelle Word ↔ Markdown. Prend en charge la conversion de formules mathématiques, la conversion bidirectionnelle des séparateurs (les trois types de lignes de séparation de Markdown vs sauts de page, sauts de section et lignes horizontales de Word), ainsi que la restauration des marker explicites `<` / `^` des tableaux Markdown en fusions rectangulaires de cellules Word. Prend en charge les formats tels que DOCX/DOC/WPS/RTF/ODT.
- **📊 Conversion de format de feuille de calcul** - Conversion bidirectionnelle Excel ↔ Markdown. Prend en charge les formats XLSX/XLS/ET/ODS/CSV/TSV, les stratégies configurables d'export des cellules fusionnées (`fill / empty / marker`) et les outils de résumé de tableau. Les modèles Markdown→XLSX prennent de nouveau en charge les champs YAML et les placeholders verticaux et horizontaux de colonne ; la restauration complète des modèles Excel, images et fusions reste un objectif de parité suivi.
- **📑 PDF et fichiers de mise en page** - Conversion PDF/XPS/OFD vers Markdown ou DOCX. Prend en charge la fusion, la division de PDF et d'autres opérations.
- **🖼️ Traitement d'image** - Prend en charge la conversion bidirectionnelle et la compression des formats JPEG/PNG/GIF/BMP/TIFF/WebP/HEIC.
- **📥 Importation d'autres formats** - Prend en charge la conversion unidirectionnelle de HTML/MHTML/ENEX/EPUB/PPTX/PPT vers Markdown.
- **🔍 Reconnaissance de texte OCR** - RapidOCR intégré pour extraire du texte à partir d'images et de PDF.
- **✏️ Correction de texte** - Vérifie les fichiers Word (.docx) et Markdown (.md) pour les fautes de frappe, la ponctuation, les symboles et les mots sensibles en fonction de dictionnaires personnalisés. Les règles peuvent être modifiées dans l'interface des paramètres.
- **📝 Système de modèles** - Mécanisme de modèle flexible prenant en charge les formats de documents et de rapports personnalisés.
- **💻 Fonctionnement en double mode** - Interface utilisateur graphique (GUI) + Interface en ligne de commande (CLI).
- **🔒 Traitement local avec protection des sorties des dépendances** - La conversion ne dépend d’aucun service en ligne. Pendant l’exécution de DocWen, son processus Python bloque DNS et IPv4/IPv6 pour les dépendances internes ; les applications Office externes conservent la politique réseau du système.
- **🔗 Fonctionnement à instance unique** - Gère automatiquement les instances de programme et prend en charge l'intégration avec le plugin Obsidian associé.

## 📸 Captures d’écran

| Lot | Markdown |
| --- | --- |
| ![Panneau lot](../assets/screenshots/batch-light.png) | ![Fenêtre principale](../assets/screenshots/main-light.png) |

| Document | Tableur |
| --- | --- |
| ![Panneau document](../assets/screenshots/conversion-document-light.png) | ![Panneau tableur](../assets/screenshots/conversion-spreadsheet-light.png) |

| Image | Fichiers de mise en page |
| --- | --- |
| ![Panneau image](../assets/screenshots/conversion-image-light.png) | ![Panneau mise en page](../assets/screenshots/conversion-layout-light.png) |

Journal des modifications : voir [CHANGELOG.md](../CHANGELOG.md)

## 🚀 Démarrage rapide

### Installation depuis le code source

**Prérequis** : Python 3.12

**Périmètre cible de la version 0.9** : Ce code source produit des paquets Windows x64 et Ubuntu
24.04 x64. Les autres distributions Linux et macOS restent des voies source/développement, non
couvertes par le paquet Ubuntu.

**Option 1 : Avec uv (recommandé)**

Installez [uv](https://docs.astral.sh/uv/getting-started/), puis :

```bash
git clone https://github.com/ZHYX91/docwen.git
cd docwen
uv sync --frozen --all-extras
```

Le code source, les tests et les builds de DocWen 0.9 prennent uniquement en charge le verrou versionné avec `uv 0.12.0` ; `pip install -e` n'est pas pris en charge.

### Lancer le programme

Pour la version packagée Windows : double-cliquez sur `DocWen.exe` pour démarrer l'interface graphique. Après installation depuis le code source :

```bash
docwen-gui  # Mode GUI
docwen      # Mode CLI
```

### Notes pour macOS

**Limitation actuelle** : Sous macOS, `convert`, `validate`, `number`, `merge` et `split` sont
actuellement indisponibles. Les notes ci-dessous documentent uniquement des dépendances facultatives
pour les expériences de développement.

**Support LibreOffice (Optionnel)**

Pour convertir des formats hérités comme `.doc` et `.xls`, installez LibreOffice :  
Téléchargement : https://www.libreoffice.org/download/

**Support des images HEIC (Optionnel)**

Pour traiter les images HEIC/HEIF :

```bash
brew install libheif
pip install pillow-heif
```

### Prérequis GUI sous Linux

**Cible packagée prise en charge** : DocWen 0.9 prend en charge la GUI et la CLI du paquet Ubuntu
24.04 x64. Ces prérequis n'étendent pas cet engagement à une autre distribution ou architecture.

- Environnement de bureau installé (GNOME, KDE, XFCE, etc.)
- L’interface graphique utilise PySide6 (Qt6) et ne dépend plus de Python Tk. Si le démarrage échoue à cause de bibliothèques système manquantes, installez les dépendances Qt indiquées par l’erreur (souvent liées à OpenGL/X11).
- Pour les serveurs headless, privilégiez l’entrée CLI `docwen` plutôt que l’interface graphique ; les builds Windows packagés fournissent aussi `DocWenCLI.exe`.

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

1.  **Lancer le programme** : Double-cliquez sur `DocWen.exe` (version Windows packagée) ou exécutez `docwen-gui`.
2.  **Importer le fichier** :
    -   Méthode 1 : Faites glisser et déposez les fichiers directement dans la fenêtre.
    -   Méthode 2 : Cliquez sur le bouton "Ajouter" dans la zone de glisser-déposer pour sélectionner des fichiers.
3.  **Sélectionner le modèle** (si conversion nécessaire) : Le panneau de modèle droit s'étend automatiquement ; sélectionnez un modèle approprié.
4.  **Configurer les options** : Cochez les options de conversion/exportation requises dans le panneau d'opération.
5.  **Exécuter l'opération** : Cliquez sur le bouton de fonction correspondant (ex : "Export MD", "Convertir en DOCX", etc.).
6.  **Voir le résultat** : La barre d’état affiche la progression et le résultat ; cliquez sur l’action « Ouvrir la sortie » à droite pour ouvrir l’emplacement de sortie.

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
| Markdown | Convertir DOCX, Convertir PDF, Correction texte |
| Tableau Excel | Export MD, Convertir PDF, Résumé tableau |
| Fichier PDF | Export MD, Fusionner, Diviser, OCR |
| Fichier Image | Conversion format, Compression, OCR |
| HTML/EPUB/PPTX etc. | Export MD |

### Interface des paramètres

Cliquez sur le bouton « Paramètres » dans l’en-tête de la zone d’opérations pour ouvrir la configuration :

Les paramètres sont organisés en onglets : **Général**, **Texte**, **Relecture**, **Document**, **Tableau**, **Image**, **Mise en page**, **Lien**, **Formatage**, **Sortie**, **Exporter**, **Journal**, **Autres**.

### Raccourcis

-   **Faire glisser fichier externe** : Faites glisser directement dans la fenêtre pour importer.
-   **Ouvrir la sortie** : Cliquez sur l’action « Ouvrir la sortie » à droite dans la barre d’état pour ouvrir l’emplacement de sortie.
-   **Clic droit élément modèle** : Ouvrir l'emplacement du fichier modèle.

---

## 🔧 Utilisation en ligne de commande

En plus de l'interface graphique, DocWen fournit une interface en ligne de commande (CLI) pour l'automatisation, les traitements par lot et les intégrations externes.

### Flux recommandé pour l'automatisation

Pour les scripts, agents ou plugins, l'ordre recommandé est le suivant :

1. `inspect <file> [--json]` : détecter d'abord la catégorie réelle du fichier, son format et les actions prises en charge.
2. `schema convert` : lire le contrat lisible par machine et les contraintes conditionnelles de `convert`.
3. `convert <file> --to <fmt> --output <path> --dry-run --json` : prévisualiser détection, normalisation et routage sans écrire de sortie.
4. `convert <file> --to <fmt> --output <path> ...` : lancer ensuite la conversion réelle.

### Exemples courants

```bash
# Paquet Windows
DocWenCLI.exe inspect document.docx --json

# Exporter le contrat de convert pour scripts / agents
DocWenCLI.exe schema convert

# Prévisualiser l'exécution sans écrire de fichiers
DocWenCLI.exe convert report.docx --to md --output report.md --extract-img --ocr --dry-run --json

# Exporter Word vers Markdown (extraction d'images + OCR)
DocWenCLI.exe convert report.docx --to md --output report.md --extract-img --ocr

# Markdown vers Word (modèle + mode de fusion titre/corps)
DocWenCLI.exe convert document.md --to docx --output document.docx --template template.docx.abaa27b964bf8589a411a696f0f5781e0917b8950067268ca15c6941668e39e9 --heading-merge-mode punct_required

# Contrôler le mode image et l'emplacement du texte OCR dans Markdown
DocWenCLI.exe convert report.docx --to md --output report.md --extract-img --image-mode file --ocr --ocr-placement image_md

# Consulter les capacités d’exécution et les portes de dépendance
DocWenCLI.exe doctor --json
DocWenCLI.exe resources list formats --json

# Vérification de documents
DocWenCLI.exe validate document.docx --check typo --check punct
DocWenCLI.exe validate input.md --check typo --check punct

# Depuis les sources / uv
# inspect -> schema -> dry-run -> convert
# docwen inspect document.docx --json
# docwen schema convert
# docwen convert document.docx --to md --output document.md --dry-run --json
# docwen convert document.docx --to md --output document.md
```

### Commandes et options courantes

Le tableau ci-dessous ne liste que les commandes les plus courantes. Pour la surface complete des commandes, utilisez `docwen --help` (code source / uv) ou `DocWenCLI --help` (version empaquetee).

| Commande / option | Description |
| --- | --- |
| `convert <file> --to <fmt> --output <path>` | Point d'entrée unifié pour les conversions. |
| `convert <file> --to <fmt> --output <path> --dry-run --json` | Prévisualise détection, normalisation, routage et options effectives sans exécuter la conversion réelle. |
| `schema convert` | Exporte le contrat lisible par machine, les valeurs par défaut, les conditions et les clés canoniques de `convert`. |
| `validate <file> --check ...` | Correction documentaire (`typo/punct/symbol/sensitive/all/none`). Utilisez `--json` pour l'enveloppe CLI ; `--report` est un chemin de fichier de rapport facultatif. |
| `inspect <file> [--json]` | Inspecte la catégorie/le format du fichier, les actions recommandées et les avertissements de décalage entre extension et contenu. |
| `doctor --json` | Retourne les diagnostics avec le résumé des capacités d’exécution et les portes de dépendance. |
| `resources list formats --json` | Liste les formats cibles par catégorie source avec des résumés des dépendances à l’exécution et des limitations. |
| `resources list templates` | Liste les modèles disponibles. |
| `resources list numbering-schemes` | Liste les schémas de numérotation disponibles. |
| `--template <id>` | ID canonique exact renvoyé par `resources list templates` ; les noms affichés, noms de fichier et chemins sont rejetés. Les ID DOCX s'appliquent à `docx/doc/odt/rtf/wps/pdf`, les ID XLSX à `xlsx/xls/ods/csv`. |
| `--extract-img` / `--no-extract-img` / `--ocr` | Extraction d'images et OCR pour `convert --to md`. |
| `--image-mode file|base64` | Contrôle la façon dont les images sont émises pendant l'export Markdown. |
| `--ocr-placement image_md|main_md` | Contrôle si le texte OCR est écrit dans le Markdown associé à l'image ou dans le Markdown principal. |
| `--heading-merge-mode punct_required|always|never` | Contrôle la stratégie de fusion titre + corps pour `convert --to docx`. |
| `--optimization <id>` | Active explicitement un profil d'optimisation (voir `resources list optimizations`). |
| `batch convert|validate ... --jobs <n> [--continue-on-error]` | Contrôles de traitement par lot. |
| `--json` / `--quiet` / `--timing` | Sortie structurée, logs réduits et données de durée pour scripts ou plugins. |

En mode `punct_required`, la liste exacte par défaut est `。：！？.:!?`. Elle peut être modifiée dans les paramètres de mise en forme ; une valeur vide désactive la fusion dans ce mode. Les virgules, points-virgules, virgules d'énumération, tirets et points de suspension sont exclus par défaut.


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

**Paragraphes mixtes** : Lorsqu'un sous-titre doit être mélangé avec le corps du texte dans le même paragraphe (mode par défaut : « Ponctuation requise »), les conditions suivantes doivent être remplies :
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
- Par défaut (mode « Ponctuation requise »), si le sous-titre ne se termine pas par une ponctuation de fin, il ne fusionne pas avec la ligne suivante même sans ligne vide.
- Vous pouvez modifier ce comportement dans Paramètres → Formatage → « MarkDown vers document » → « Heading + body merge mode ».

### Conversion bidirectionnelle des séparateurs

Prend en charge la conversion bidirectionnelle entre les séparateurs Markdown et les sauts de page/sauts de section/lignes horizontales Word :

-   **DOCX → MD** : Les sauts de page, sauts de section et lignes horizontales Word sont automatiquement convertis en séparateurs Markdown.
-   **MD → DOCX** : Markdown `---`, `***`, `___` sont automatiquement convertis en éléments Word correspondants.
-   **Configurable** : Les relations de mappage spécifiques peuvent être personnalisées dans l’interface des paramètres.

### Listes de tâches

Prend en charge la conversion bidirectionnelle des listes de tâches GFM :

```markdown
- [ ] À faire
- [x] Terminé
```

-   **MD → DOCX** : Rendu sous forme de liste à puces avec préfixe texte `☐` / `☑`.
-   **DOCX → MD** : Convertit les éléments de liste commençant par `☐` / `☑` / `☒` en `- [ ]` / `- [x]`.
-   **Note sur les polices** : `☐`/`☑` peuvent ne pas s’afficher avec certaines polices. Si nécessaire, utilisez des polices comme « Segoe UI Symbol » dans votre modèle Word.

### Insertion d’images et taille

Prend en charge l’insertion d’images Obsidian/Wiki et Markdown standard, avec taille optionnelle (px) :

```markdown
![[image.png]]
![[image.png|300]]
![[image.png\|300]]
![alt](image.png =300x200)
![alt](image.png =300x)
![alt|300](image.png)
```

- Sans taille : taille d’origine, limitée par la largeur disponible (page/cellule)
- Avec taille : agrandissement autorisé, toujours limité par la largeur disponible
- Paragraphe uniquement image : utilise le style de paragraphe « Image » (centré, interligne simple)

### Gestion des liens

Prend en charge les liens cliquables en Markdown -> DOCX :

```markdown
[Docwen](https://example.com)
[[Target]]
[[Target|Open target]]
<https://example.com>
<user@example.com>
```

- Les liens Markdown et Wiki sont convertis par défaut en hyperliens Word
- Les liens Wiki sont résolus en liens locaux `file:///` lorsque le fichier cible est trouvé
- Les autolinks entre chevrons prennent en charge `https://...` et les e-mails `mailto:...`
- La création automatique de liens pour les URL nues est évaluée par requête pour Markdown -> DOCX, désactivée par défaut et activée par `[non_embed_links].auto_link_bare_url` dans `configs/link.toml`
- Markdown -> XLSX ne génère pas de placeholders d'hyperliens DOCX et conserve la syntaxe d'origine

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
| **Optimisation avancée des champs** | Si coché, extrait des métadonnées structurées plus riches ; sinon, utilise le mode simplifié avec uniquement le titre et le sous-titre. |
| **Nettoyer numéros sous-titres** | Si coché, supprime les numéros avant les sous-titres (ex: "一、", "（一）", "1.", etc.) et les convertit en texte de titre pur. |
| **Ajouter numéros sous-titres** | Si coché, ajoute automatiquement des numéros en fonction des niveaux d'en-tête (le schéma de numérotation peut être configuré dans les paramètres). |

Remarque : DOCX -> MD restaure désormais aussi la numérotation multiniveau liée dans numbering.xml via les styles de paragraphe (pStyle). Les préfixes de titre créés avec les listes multiniveaux de Word/WPS, comme "一、", "（一）", "1．", "（1）" et "①", sont donc conservés en mode simplifié comme en mode avancé des champs ; le niveau de titre reste correctement détecté lorsque l’option "Nettoyer numéros sous-titres" est activée.

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
2.  **Markdown vers Excel** : Les tableaux Markdown peuvent être exportés en XLSX. Les modèles XLSX prennent en charge les champs YAML, les placeholders de colonne et d'image, ainsi que les cellules fusionnées ou protégées.

**Formats pris en charge** :
-   `.xlsx` - Document Excel standard.
-   `.xls` - Automatiquement converti en XLSX pour le traitement.
-   `.et` - Feuille de calcul WPS automatiquement convertie.
-   `.csv` - Tableau texte CSV.
-   `.tsv` - Tableau TSV séparé par des tabulations.


### Fonction de correction de texte

Le programme fournit quatre règles de correction personnalisables :

1.  **Vérification de l'appariement de la ponctuation** - Détecte si la ponctuation par paire comme les parenthèses et les guillemets correspond.
2.  **Correction de symboles** - Détecte l'utilisation mixte de la ponctuation chinoise et anglaise.
3.  **Vérification des fautes de frappe** - Vérifie les fautes de frappe courantes en fonction d'un dictionnaire personnalisé.
4.  **Détection de mots sensibles** - Détecte les mots sensibles en fonction d'un dictionnaire personnalisé.

**Dictionnaires personnalisés** : Modifiez visuellement les dictionnaires de fautes de frappe et de mots sensibles dans l'interface "Paramètres".

**Utilisation** :
1.  Faites glisser le document Word ou le fichier Markdown à vérifier dans le programme.
2.  Cochez les règles de correction requises.
3.  Cliquez sur le bouton "Correction de texte".
4.  Les résultats de la correction sont affichés sous forme de commentaires dans le document. Pour les fichiers Markdown, un rapport JSON est généré.

Note (rapport JSON de correction Markdown) :
- Moteur : `text_rules` + adaptateur Markdown `md_spell`
- Sortie : l'entrée CLI actuelle de correction est `validate` ; utilisez `--json` pour l'enveloppe CLI. `--report` est un chemin de fichier de rapport facultatif.

- Différent de `--json` (JSON enveloppe CLI)

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

#### Espaces réservés de modèle Excel (objectif de parité hérité)

Les modèles XLSX prennent en charge les champs YAML, les placeholders verticaux `{{↓champ}}` et horizontaux `{{→champ}}`, les placeholders d'image, ainsi que les cellules fusionnées ou protégées.

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

**Traitement des cellules fusionnées** :

- Markdown -> Excel conserve les merged ranges d'origine du modèle.
- Dans les zones de modèle en colonnes connues composées de placeholders contigus `{{↓NomChamp}}`, le programme peut restaurer des fusions rectangulaires à partir de marqueurs explicites `<` / `^` dans les tableaux Markdown.
- Seules les cellules dont le contenu, une fois les espaces de début et de fin supprimés, est exactement `<` ou `^` participent à la détection des fusions ; `\<` et `\^` restent du texte littéral.
- Les rectangles invalides ou les conflits avec les merged ranges existants du modèle sont rétrogradés en texte normal avec un avertissement, au lieu d'écraser de force la structure du modèle.

**Fusion de données multi-tableaux** : S'il y a plusieurs tableaux dans Markdown utilisant le même nom d'en-tête, les données seront fusionnées dans l'ordre et remplies séquentiellement.

## 🔌 Plugin Obsidian

Un plugin Obsidian compagnon est publié séparément et fonctionne avec le convertisseur :

### Fonctionnalités principales

-   **🚀 Lancement en un clic** - Icône de la barre latérale pour lancer rapidement le convertisseur.
-   **📂 Transfert automatique** - Passe automatiquement le chemin du fichier actuellement ouvert.
-   **🔄 Gestion d'instance unique** - Envoie automatiquement le fichier si le programme est déjà en cours d'exécution, pas besoin de redémarrer.
-   **🔒 Contrôle local borné** - Utilise des requêtes typées `status`, `open` et `activate`, sans recherche de processus par nom ni fichiers de commande ou d'état.

### Principe de fonctionnement

Le transport runtime/control de DocWen Core utilise un canal nommé Windows ou un socket AF_UNIX sous
Linux/macOS. Un verrou de fichier établit uniquement la propriété de l'instance unique ; aucun fichier
ne transporte les commandes de contrôle. Cela décrit uniquement la capacité du Core. DocWen Assistant
2.0 reste limité au bureau Windows et ne dispose d'aucune recette combinée sous Linux/macOS.

1.  **Premier clic** → Lancer le convertisseur et passer le fichier actuel.
2.  **Cliquer à nouveau (Avec fichier)** → Remplacer par le nouveau fichier (Mode fichier unique).
3.  **Cliquer à nouveau (Sans fichier)** → Activer la fenêtre du convertisseur.

### Installation

DocWen Assistant 2.0 utilise DocWen Machine Protocol v1 et l'unique contrat Artifact Bundle v2. La version du code
source ne prouve pas sa publication ; installez uniquement une version numérique qui identifie explicitement une
version publiée et compatible de DocWen.

## 🔌 OpenClaw (Plugin + Skill)

OpenClaw 2.0 utilise DocWen Machine Protocol v1 et l'unique contrat Artifact Bundle v2. La version du code source ne
prouve pas sa publication ; consultez la page de la version numérique et ne l'installez qu'après la réussite de son
contrôle de publication immuable.

## ❓ Foire aux questions (FAQ)

### Que faire si la conversion échoue ?

-   Vérifiez si le fichier est occupé par un autre programme.
-   Confirmez que le format du fichier est correct.
-   Consultez dans les paramètres le champ « Chemin réel actuel du fichier journal » ou vérifiez les journaux d'erreurs dans le répertoire de journaux utilisateur du système ; si la vérification du paquet utilise `DOCWEN_LOG_DIR`, consultez à la place le répertoire surchargé.

### Le modèle ne s'affiche pas ?

-   Confirmez que les fichiers modèles sont dans le répertoire `templates/`.
-   Vérifiez si le fichier modèle est corrompu.
-   Redémarrez le programme pour recharger les modèles.

### La fonction de correction ne fonctionne pas ?

-   Confirmez que le document est au format .docx ou .md.
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

## 🔒 Fonctionnalites de securite

-   **Fonctionnement completement local** : Le traitement se fait localement par defaut et ne depend pas de services en ligne.
-   **Protection des sorties des dépendances** : Les entrées GUI/CLI prises en charge activent un garde d’audit CPython pendant toute la durée du processus Python principal. Il bloque toute résolution DNS/de noms et les opérations AF_INET/AF_INET6 `bind`, `connect`, `connect_ex`, `sendto` et `sendmsg`, tout en conservant les canaux nommés Windows et les sockets de domaine Unix.
-   **Limite explicite** : Les processus lancés séparément, dont Office/WPS/LibreOffice et l’assistant Office, ne sont pas gérés. Il s’agit d’une défense contre les connexions accidentelles des dépendances, pas d’un bac à sable du système.
-   **Pas de televersement de donnees** : Par defaut, les fichiers utilisateur ne sont pas activement envoyes vers des serveurs externes.
-   **Mode de securite strict** : active par defaut ; l'application se ferme si les controles de securite centraux echouent. Voir [Troubleshooting](../maintenance/troubleshooting.md).

## 📜 Licence

Ce projet est sous licence **GNU Affero General Public License v3.0 (AGPL-3.0)**.

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

-   Ce projet utilise PyMuPDF (sous licence AGPL-3.0), donc l'ensemble du projet est également sous licence AGPL-3.0.
- La GUI actuelle peut utiliser `PySide6-Fluent-Widgets` (QFluentWidgets) sur les chemins d’hôte pris en charge ; cette dépendance suit un modèle de double licence `GPLv3 / commerciale`, tandis que ce dépôt continue d’être distribué sous AGPL.
-   Vous êtes libre d'utiliser, de modifier et de distribuer ce logiciel.
-   Si vous modifiez ce logiciel et fournissez des services sur un réseau, vous devez fournir le code source modifié aux utilisateurs.
-   Pour des informations détaillées sur la licence, veuillez consulter le fichier [LICENSE](../../LICENSE).
- Pour les mentions des composants tiers, consultez [LICENSE_THIRD_PARTY.txt](../../LICENSE_THIRD_PARTY.txt) ; le resume de distribution se trouve dans [NOTICE.txt](../../NOTICE.txt).

### Contact

-   **GitHub** : https://github.com/ZHYX91/docwen
-   **Contacter l'auteur** : zhengyx91@hotmail.com

---

**Auteur** : ZhengYX
