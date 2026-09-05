# Cahier des charges fonctionnel PADRES HTML / JavaScript

**Version : 3.5**
**Date de mise à jour : 15/08/2026**

## Historique des modifications

| Version | Date | Modification |
|---------|------|-------------|
| 1.0 | — | Version initiale |
| 2.0 | 09/08/2026 | Ajout page Débits réservés, logo SICASIL, validation fichier, palettes couleurs, seuils d'alerte |
| 3.0 | 13/08/2026 | Ajout page Production, nettoyage complet des libellés, graphe jour de pointe, refonte page App, graphique production en lignes avec aires |
| 3.1 | 14/08/2026 | Filtre d'outliers sur les volumes de production (plafond 200 000 m³/j par usine) pour exclure les valeurs aberrantes type index cumulé |
| 3.2 | 14/08/2026 | Affichage des bornes de période à droite du sélecteur ; KPI « Mesure la plus récente » = dernière date du fichier |
| 3.3 | 14/08/2026 | Filtrage des valeurs nulles sur les « Débit station hydrométrique » (page Débits réservés) |
| 3.4 | 15/08/2026 | Correctif retaillage des graphiques, filtre « 6 derniers mois », affichage mobile avec menu hamburger |
| 3.5 | 15/08/2026 | Glisser-déposer de fichier dans la page App |

---

## 1. Objectif et périmètre

Créer une application web PADRES autonome qui se lance en ouvrant simplement `index.html` dans un navigateur.

L'application doit permettre à un utilisateur non technique de sélectionner un fichier Excel PADRES local et de visualiser les mêmes indicateurs que dans l'application Streamlit actuelle.

Toutes les données doivent rester sur le poste utilisateur. Aucune donnée ne doit être envoyée à un serveur distant.

L'application doit reproduire au mieux le comportement, le style visuel, les calculs, les filtres et les graphes illustrés par les captures de référence.

## 2. Contraintes techniques

- Un seul fichier `index.html` est acceptable, éventuellement accompagné de fichiers CSS/JS si nécessaire.
- Aucun backend, aucun serveur local, aucun Python, aucun NodeJS, aucun framework lourd.
- Bibliothèques autorisées : **SheetJS** pour lire Excel et **Plotly.js** pour les graphiques.
- Le fichier Excel doit être chargé localement :
  - En essayant en premier lieu de recharger le dernier fichier utilisé (via IndexedDB)
  - ou via un `input type=file` si c'est la première utilisation ou si le dernier fichier utilisé n'existe plus
- Le rendu doit fonctionner sur les navigateurs modernes, notamment Microsoft Edge, Chrome et Safari.

## 3. Architecture de l'application

Navigation latérale avec **six pages** :

1. 📊 **Tableau de bord**
2. 💧 **Ressources**
3. 🏭 **Production**
4. 🚰 **Ventes en gros**
5. 🚱 **Débits réservés**
6. ⚙️ **App**

Le menu doit rester visible sur toutes les pages.

La page Tableau de bord doit être la page d'accueil par défaut.

La page App est dédiée au chargement du fichier Excel et au diagnostic.

Les graphes et données ne doivent être visibles qu'après chargement réussi du fichier Excel.

### 3.1 Logo SICASIL

- Le logo SICASIL (https://sicasil.com/wp-content/uploads/2020/09/logo-sicasil-2.jpg) est affiché en haut de la barre latérale, à la place d'un texte.
- Le fond de la zone du logo est gris clair (identique au fond de la barre latérale).
- La propriété CSS `mix-blend-mode: multiply` est appliquée au logo pour que les parties blanches deviennent transparentes et s'affichent naturellement sur le fond gris.
- La zone de sélection des pages reste en gris clair.
- Dimensions contraintes : `max-width: 100%`, `max-height: 60px`, `object-fit: contain`.

## 4. Chargement du fichier Excel et état initial

- Avant chargement, essayer de recharger le fichier utilisé lors de la dernière utilisation (stocké dans IndexedDB).
- En cas d'échec ou si c'est le premier lancement, les pages Tableau de bord, Ressources, Production, Ventes en gros et Débits réservés doivent afficher un état vide : « Aucun fichier chargé. Veuillez sélectionner un fichier Excel dans la page App. »
- La page App doit afficher un composant de sélection de fichier acceptant `.xlsx`, `.xlsm` et `.xls`.
- Le texte de sélection doit préciser : **« du type PADRES données historiques »** pour éviter que l'utilisateur essaie de charger un autre type de fichier.
- La feuille `data` doit être utilisée en priorité. Si elle n'existe pas, utiliser la première feuille du classeur.

### 4.1 Validation du fichier

**Conditions de validation :**

| Condition | Seuil |
|-----------|-------|
| Colonnes mappées | ≥ 40 |
| Lignes de données | ≥ 1000 |

**En cas d'échec de validation :**
- Le fichier est rejeté.
- Un message d'erreur de format s'affiche dans la page App : « ❌ Erreur de format : le fichier sélectionné ne semble pas être un fichier PADRES données historiques valide ».
- L'application se comporte comme si aucun fichier n'avait été chargé (état vide sur toutes les pages).

### 4.2 Diagnostic de chargement

Une fois le fichier chargé, la section « 📋 Diagnostic de chargement » affiche :
- Le **nom du fichier**
- Le **chemin / dossier** complet (si disponible dans le navigateur, sinon le nom du fichier)
- Le nombre de **lignes conservées**
- Le nombre de **colonnes**
- La **période détectée**

Le message « ✅ Fichier chargé avec succès » reste **en dehors** des sections repliables, toujours visible.

### 4.3 Gestion du rechargement automatique (IndexedDB)

- Le fichier est stocké comme un **`Blob`** dans IndexedDB (format le plus fiable pour le binaire sur tous les navigateurs, y compris Safari en `file://`).
- L'ArrayBuffer est **copié** (`slice(0)`) avant d'être passé à `XLSX.read()` car cette bibliothèque **détache** l'ArrayBuffer.
- Au rechargement, vérifier que le Blob a une taille > 0 avant de l'utiliser.
- Si le cache est corrompu ou vide, il est **nettoyé automatiquement** et l'état vide est affiché.
- Le chemin complet du fichier est stocké dans `localStorage` pour être réaffiché après rechargement.

```js
const workbook = XLSX.read(arrayBuffer, { type: "array" });
const sheetName = workbook.SheetNames.includes("data") ? "data" : workbook.SheetNames[0];
const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: true });
```

## 5. Nettoyage, filtrage et structuration des données

- Ligne 2 du fichier Excel : en-têtes métier.
- Ligne 5 du fichier Excel : début des données.
- La première colonne est la date ; si les cellules contiennent aussi des heures, seule la date doit être considérée.
- Créer `DateJour`, `Année`, `Mois` et `JourAnnée` après conversion des dates.
- **Stocker le timestamp numérique** (`time`) en plus de l'objet Date pour les comparaisons fiables sur tous les navigateurs.
- Supprimer les lignes totalement vides.
- Conserver uniquement les lignes dont les deux premières colonnes métier après la date sont renseignées.
- Supprimer les lignes dont toutes les colonnes métier sont vides.
- Supprimer les lignes dont toutes les valeurs métier sont égales à zéro.
- Supprimer les dates invalides.
- Trier les données par date croissante (sur le timestamp numérique).

### 5.1 Nettoyage des noms de colonnes

Supprimer les préfixes et marqueurs suivants :
- `TONI` en début de libellé
- `(C)`
- `(SM)`
- `(Via Idx)`, `(ViaIdx)`, `(Via Idx+)`, `(viaDébit)`, `(Via Index)`, `(ORAi1)`
- Réduire les espaces multiples à un seul espace

```js
function cleanLabel(label) {
  return String(label)
    .replace(/^TONI\s+/i, "")
    .replace(/\(C\)\s*/g, "")
    .replace(/\(SM\)\s*/g, "")
    .replace(/\(Via Idx\)\s*/gi, "")
    .replace(/\(ViaIdx\)\s*/gi, "")
    .replace(/\(Via Idx\+\)\s*/gi, "")
    .replace(/\(viaDébit\)\s*/gi, "")
    .replace(/\(Via Index\)\s*/gi, "")
    .replace(/\(ORAi1\)\s*/gi, "")
    .replace(/\s+/g, " ")
    .trim();
}
```

La colonne 0 (date) est forcée au libellé « Date ».

## 6. Mapping des colonnes et calculs métier PADRES

Toutes les valeurs sont exprimées en mètres cubes (m³), sauf indication contraire.

### Ressources naturelles

| Colonne |
|---------|
| Volume jour Total Barrage + Canal St Cezaire |
| Volume jour Total Foux de St Cezaire |
| Volume jour Total Sources de Gréolières |
| Volume jour Total source de Bramafan |
| Côte St-Cassien journalière (valeur en m) |

### Réglementation (droits d'eau)

| Colonne |
|---------|
| Droit d'eau journalier prise en rivière de la Siagne à St Cézaire |
| Droit d'eau journalier prise en rivière de la Foux de St Cézaire |
| Droit d'eau journalier prise en Siagne des Jacourets |
| Droit d'eau journalier prise en Siagne des Veyans |
| Droit d'eau journalier prise en Siagne de l APIE |
| Droit d'eau journaliers des PDR Vallée de la Siagne |
| Droit d'eau journalier des sources de Gréolières et Bramafan |
| Droit d'eau journalier prise en rivière du Loup de Bramafan |

### Ressources disponibles

| Colonne |
|---------|
| Volume jour disponible prise d'eau de St Cézaire |
| Volume jour disponible Sources de Gréolières |
| Volume jour Disponible prise en Siagne de l APIE |
| Volume jour Disponible PDR basse vallée de la Siagne |
| Disponibilité selon période estivale de la ressource des Veyans |

### Points de contrôles de débits

| Colonne | Unité |
|---------|-------|
| Débit minimum journalier au pont de Pégomas | l/s |
| Débit moyen journalier au pont de Pégomas | l/s |
| Débit Station hydrométrique La Siagne à Pégomas | l/s |
| Débit Station hydrométrique La Siagne à Callian (Ajustadoux) | l/s |

### Prélèvements

| Colonne |
|---------|
| Volume jour Prise en rivière St Cezaire |
| Volume jour source la Foux de St Cezaire |
| Volume jour Prise en rivière Les Veyans |
| Volume jour pompage eau brute DN1000 Apie |
| Volume jour PDR1 |
| Volume jour PDR2 |
| Volume jour PDR7 |
| Volume jour source de Bramafan |
| Volume jour pompage jacourets (eau brute) |

### Capacité de production

| Colonne |
|---------|
| Capacité de pompage journalier prise en Siagne de l APIE |
| Capacité de pompage journalier des PDR Vallée de la Siagne |

### Volumes produits

| Colonne |
|---------|
| Volume jour entrée usine Nartassier |
| Volume livré journalier Usine de St Jacques |
| Volume journalier produit par Chateauneuf |
| Volume jour refoulement ET Apie |
| Production totale Basse Vallée de la Siagne Auribeau-Pégomas |

**Total production SICASIL** = somme des 5 colonnes ci-dessus.

#### Filtre d'outliers sur les volumes de production

Certaines cellules Excel peuvent contenir des valeurs matériellement impossibles pour un volume journalier (ex. un **index cumulé** saisi par erreur à la place d'un volume jour). Exemple constaté : le 29/12/2016, Nartassier = 57 182 760 m³.

Pour chaque usine, toute valeur de volume produit hors plage est **ignorée** (traitée comme 0) :

| Règle | Seuil |
|-------|-------|
| Valeur négative | exclue |
| Valeur > **200 000 m³/j** par usine | exclue (outlier) |

Le plafond 200 000 m³/j est largement supérieur aux capacités théoriques maximales (~57 000 m³/j) tout en restant réaliste pour un volume journalier total multi-usines.

Ce filtre s'applique :
- au calcul du **Total production SICASIL** (`getProduction`)
- aux **KPI** et graphiques du Tableau de bord qui utilisent ce total
- aux **barres « Volume produit »** des graphiques par usine (page Production)
- au graphe **Production du jour de pointe**

### Distribution

| Colonne |
|---------|
| Besoins propres SICASIL |
| Total des Achats d'Eau journalier du SIEF (2018->) |
| Volume Jour Vente En Gros Valbonne |
| Volume jour Cap Roux - Théoule vers Trayas Var |
| Volume jour pompage jacourets (eau brute) |

### 6.1 Compléments disponibles

On appelle **C** le complément disponible à calculer :

- **D** = Somme des volumes des ressources naturelles disponibles
- **R** = débit réservé
- **J** = droits d'eau journalier
- **P** = Somme des volumes prélevés

Volume réellement disponible : **V = MIN (MAX (0 ; D - R) ; J)**

Complément disponible : **C = V - P**

**Règles complémentaires :**
- Si aucun droit journalier J n'existe, considérer J infini.
- **Veyans** : disponibilité D = 30 000 m³/j du 15 juillet au 14 octobre, sinon 0.
- **PDR** : disponibilité D = 80 000 m³/j du 1er avril au 14 octobre, sinon 43 200 m³/j.
- **APIE** : C = MAX(0 ; D - P), sans limitation J.
- **Loup** (en aval de Gréolières + Bramafan) : R = 310 l/s du 16 octobre au 15 juillet, et 150 l/s du 16 juillet au 15 octobre.
- **Total Distribution SICASIL** = TOTAL VEG Sicasil + Besoins Propres SICASIL.

### 6.2 Seuils d'alerte

**Pont de Pégomas :**

| Niveau | Seuil (l/s) |
|--------|-------------|
| Alerte (jaune) | 800 |
| Alerte renforcée (orange) | 550 |
| Crise (rouge) | 300 |

**Ajustadoux (Callian) :**

| Niveau | Seuil (l/s) |
|--------|-------------|
| Alerte (jaune) | 700 |
| Alerte renforcée (orange) | 550 |
| Crise (rouge) | 400 |

### 6.3 Capacités théoriques des usines

| Usine | Capacité théorique (m³/jour) |
|-------|------------------------------|
| Chateauneuf | Inconnue |
| Nartassier | 50 000 |
| St Jacques | 40 000 |
| Apié | 57 000 |
| PDR | Colonne « Capacité de pompage journalier des PDR Vallée de la Siagne » |

## 7. Filtre de période global

Le filtre de période est partagé entre Tableau de bord, Ressources, Production, Ventes en gros et Débits réservés.

**Valeurs possibles :**
- Dernier jour
- 7 derniers jours
- 30 derniers jours (défaut)
- 60 derniers jours
- 6 derniers mois (période glissante)
- Année en cours
- 12 derniers mois
- Année précédente
- Personnalisé

La valeur par défaut est **30 derniers jours**.

Le choix doit rester mémorisé lors du changement de page.

Le mode Personnalisé doit afficher Date début et Date fin, avec mémorisation des valeurs. Si les dates ne sont pas définies, les bornes des données sont utilisées par défaut.

### 7.1 Affichage des bornes de période

À droite du sélecteur de période (dans la barre de filtre, **pas** dans la liste déroulante), afficher les deux dates correspondant à la sélection, sous la forme :

```
➜ du <date début> au <date fin>
```

Exemples :
- Année en cours → `➜ du 1er janvier 2026 au 16 juillet 2026`
- 30 derniers jours → `➜ du 17 juin 2026 au 16 juillet 2026`

**Format des dates** : long français (`1er` pour le 1er du mois, mois en toutes lettres, année sur 4 chiffres).

Cet affichage doit :
- apparaître sur **toutes les pages** concernées par le filtre
- être mis à jour dès le **premier chargement** du fichier (pas seulement après un changement de filtre)
- se mettre à jour à chaque changement de période ou de dates personnalisées

Les graphiques suivants **ne dépendent pas** de ce filtre global :
- Le lac de Saint-Cassien (utilise tout l'historique)
- La production du jour de pointe (utilise tout l'historique)

### 7.2 Filtre « 6 derniers mois »

Le filtre **« 6 derniers mois »** (période glissante) est disponible entre « 60 derniers jours » et « Année en cours » dans tous les sélecteurs de période.

Il est calculé de manière cohérente avec « 12 derniers mois » :
```
start = aujourd'hui - 6 mois
start = start + 1 jour   (décalage comme pour 12 mois)
end   = aujourd'hui
```

La valeur `last_6m` est gérée dans les 5 sélecteurs (`global-period`, `-2`, `-3`, `-4`, `-5`) et dans `getPeriodRange()`.

## 17. Rendu sur téléphone portable

### 17.1 Menu hamburger

- Le bouton hamburger **« ☰ »** est **caché par défaut** sur ordinateur (≥ 768 px).
- Sur **smartphone** (< 768 px), il s'affiche en haut à gauche et ouvre la sidebar en **panneau coulissant**.
- Un **overlay sombre** couvre le contenu quand le menu est ouvert ; cliquer dessus le ferme.
- Choisir une page ferme automatiquement le menu.
- Sur les écrans entre **768 et 900 px**, le comportement existant (sidebar réduite en icônes) est conservé.

### 17.2 Règles responsive mobile (< 768 px)

- Le contenu principal occupe **toute la largeur** (pas de barre latérale).
- Les **cards KPI**, **filtres** et **graphes** s'empilent verticalement.
- Les **inputs/selects** passent à `font-size: 16px` pour éviter le zoom automatique iOS.
- Les **graphes** sont moins hauts (`320 px`) pour tenir à l'écran.
- Le **tableau des ressources** défile horizontalement si nécessaire.
- Le **titre** principal est réduit en taille.

### 17.3 Correctif du retaillage des graphiques

Un utilitaire `resizeActiveCharts()` appelle `Plotly.Plots.resize()` sur tous les graphiques **visibles** (`.chart-box`, `.chart-box-sm` de la page active). Il est déclenché :
- à la fin de `showPage()` (après `requestAnimationFrame`)
- lors d'un `resize` de la fenêtre

Cela corrige les graphiques **mal retaillés** (trop petits avec espace blanc, ou débordants) qui apparaissaient après un changement de page, un changement de taille de fenêtre ou un rechargement de fichier.

## 8. Page Tableau de bord

- Titre principal : « 💧 PADRES - Tableau de bord SICASIL »
- Section « 📈 Production »

### KPI

| KPI | Affichage |
|-----|-----------|
| Mesure la plus récente | **Date** en valeur principale (en gras) — dernière date **du fichier** (indépendante du filtre de période) |
| Production totale | Total en m³ + moyenne m³/j entre parenthèses (sur la période filtrée) |
| Ventes en gros | Total en m³ + moyenne m³/j entre parenthèses (sur la période filtrée) |

### Graphique Production totale et ventes en gros

| Série | Couleur | Style |
|-------|---------|-------|
| Total production SICASIL | #1f77b4 (bleu) | Ligne + aire colorée semi-transparente (rgba(31,119,180,0.25)) |
| Total VEG SICASIL | #1f77b4 (bleu) | Ligne + aire colorée plus foncée (rgba(31,119,180,0.45)) |
| Production pointe | #d62728 (rouge) | Point taille 14, contour blanc, valeur affichée |

Le graphique utilise des **lignes avec aires colorées** sous les courbes (`fill: tozeroy`), avec les mêmes couleurs bleues que la page Production.

### 8.1 Prélèvements journaliers et compléments disponibles

Histogramme empilé par jour.

**Sélecteur : m³/jour ou Normalisé**

| Série | Couleur |
|-------|---------|
| Prélèvements | #c0392b |
| Compléments disponibles | #3498db |

En mode Normalisé, axe Y de 0 à 100 %.
Le hover doit toujours afficher le volume m³/jour et la part en %.

## 9. Page Ressources

- Section « 💧 Ressources »
- Utilise le filtre global de période

### Graphique Disponibilité détaillée des ressources

Une courbe par ressource.

| Ressource | Couleur indicative |
|-----------|-------------------|
| Barrage + Canal St-Cézaire | Bleu |
| Foux de St-Cézaire | Bleu clair |
| Sources de Gréolières | Rouge |
| Bramafan | Rose |
| Jacourets | Vert / turquoise |

### Camembert Répartition des ressources

Calculé sur la période sélectionnée + tableau des volumes par ressource.

## 10. Lac de Saint-Cassien

- La section se trouve dans la page Ressources.
- **Ne pas appliquer** le filtre global de période — utilise tout l'historique disponible.
- Afficher une courbe par année, superposée sur une année virtuelle allant du 1er janvier au 31 décembre.
- Sélecteur multi-années ; toutes les années sont sélectionnées par défaut ; l'utilisateur peut déselectionner et reselectionner des années à sa guise (les cases restent visibles même décochées).
- **Palette de couleurs** : chaque année affichée a une couleur distincte (palette de 15 couleurs), stable indépendamment de la sélection.
- L'année la plus récente est en ligne épaisse (largeur 3), les autres plus fines (largeur 1) et semi-transparentes (opacité 0.4).

### 10.1 Formule cote vers volume

Utiliser exclusivement le polynôme ordre 3 suivant. Le volume V(x) est en millions de m³ et x est la cote en mètres :

```
V(x) = 0.001586592054 × x³ - 0.6006702028 × x² + 77.42098071 × x - 3415.762614
volume_m3 = max(0, V(x) × 1_000_000)
```

### 10.2 Lignes de référence du lac

| Référence | Valeur |
|-----------|--------|
| Retenue normale | Cote 147,35 m → volume, ligne pointillée libellée « Retenue normale » |
| Minimum exploitation | Cote 138,50 m → volume, ligne pointillée libellée « Minimum exploitation » |
| 20 000 000 m³ | Ligne horizontale pointillée sans label |
| 1er juillet | Ligne verticale pointillée avec libellé « 1er juillet » |

## 11. Page Production

- Section « 🏭 Production »
- Positionnée entre « Ressources » et « Ventes en gros » dans le menu
- Utilise le filtre global de période

### 11.1 Graphiques par usine

Un graphique par usine (Chateauneuf, Nartassier, St Jacques, Apié, PDR) :

| Donnée | Type | Couleur |
|--------|------|---------|
| Volume produit | Barres | Bleu #1f77b4 |
| Capacité réelle | Ligne pointillée | Rouge #d62728 |
| Capacité théorique | Ligne pointillée | Vert #2ca02c |

Les trois données sont sur la même échelle, unité : m³/jour.

**Capacités théoriques :**
- Chateauneuf : inconnue
- Nartassier : 50 000 m³/jour
- St Jacques : 40 000 m³/jour
- Apié : 57 000 m³/jour
- PDR : colonne « Capacité de pompage journalier des PDR Vallée de la Siagne »

### 11.2 Production du jour de pointe

- **Indépendant du filtre de date** (utilise tout l'historique)
- **X** : années
- **Y** : barres bleues, une seule valeur par année = **valeur maximale** de volume journalier produit (Total production SICASIL) de cette année
- **Étiquettes sur les barres** :
  - **Volume** : affiché horizontalement **au-dessus** de chaque barre (valeur seule, sans unité « m³ », ex. `146 761`)
  - **Date** : affichée horizontalement **à l'intérieur** de chaque barre, sous la forme `dd/mm` uniquement (sans l'année, déjà indiquée sur l'axe X)
- **Hover** : volume + date complète (ex. « 146 761 m³ — 12/08/2016 »)
- Les volumes sont calculés **après application du filtre d'outliers** (plafond 200 000 m³/j par usine), afin d'éviter qu'une valeur aberrante (ex. index cumulé) ne fausse le maximum annuel

## 12. Page Ventes en gros

- Section « 🚰 Ventes en gros détaillées »
- Utilise le filtre global de période

### Graphique d'évolution

| Flux | Couleur |
|------|---------|
| Total VEG SICASIL | Bleu foncé |
| SIEF | Bleu |
| Valbonne | Rouge |
| Cap Roux / Théoule / Trayas | Rose / orange |
| Pointe VEG | Point rouge taille 14, valeur affichée |

### Camembert de répartition des ventes sur la période

Mêmes couleurs que le graphique d'évolution.

## 13. Page Débits réservés

- Section « 🚱 Débits réservés »
- Positionnée entre « Ventes en gros » et « App » dans le menu
- Utilise le filtre global de période

### Graphique 1 : Pont de Pégomas

Affiche les 3 valeurs de débit disponibles à Pégomas :
- Débit minimum journalier au pont de Pégomas (l/s)
- Débit moyen journalier au pont de Pégomas (l/s)
- Débit Station hydrométrique La Siagne à Pégomas (l/s)

Affiche également le **débit réservé** :
- 310 l/s du 16 octobre au 15 juillet
- 150 l/s du 16 juillet au 15 octobre

**Lignes horizontales pointillées des seuils d'alerte :**

| Niveau | Couleur | Seuil (l/s) |
|--------|---------|-------------|
| Alerte | Jaune | 800 |
| Alerte renforcée | Orange | 550 |
| Crise | Rouge | 300 |

### Graphique 2 : La Siagne à Callian (Ajustadoux)

Affiche le débit de la station hydrométrique.

**Lignes horizontales pointillées des seuils d'alerte :**

| Niveau | Couleur | Seuil (l/s) |
|--------|---------|-------------|
| Alerte | Jaune | 700 |
| Alerte renforcée | Orange | 550 |
| Crise | Rouge | 400 |

### 13.1 Filtrage des valeurs nulles (débits station hydrométrique)

Pour les séries **« Débit station hydrométrique »** des deux graphiques (Pont de Pégomas et Ajustadoux/Callian), les valeurs **0 ou vides** correspondent à une **absence de mesure** et non à un débit réel nul.

Ces valeurs sont donc **exclues de l'affichage** :

- Elles sont converties en `null` (aucun point tracé) dans la série.
- `connectgaps` est désactivé (`false`) : aucune interpolation entre deux points séparés par un trou.
- La courbe **s'arrête** et ne descend **pas artificiellement à zéro** en fin de période.
- Les autres séries des graphiques (Débit minimum journalier, Débit moyen journalier, Débit réservé, seuils d'alerte) ne sont pas concernées par ce filtre.

## 14. Page App

- Permettre le chargement du fichier Excel.
- Texte de sélection : « Cliquez pour sélectionner un fichier, ou faites glisser ici un fichier Excel du type « PADRES données historiques » ».
- **Glisser-déposer** : l'utilisateur peut déposer un fichier directement dans le cadre principal (événements `dragover`/`dragleave`/`drop`). Le cadre s'illumine (bordure verte) au survol d'un fichier.
- Une **icône animée** (👇) indique visuellement que le glisser-déposer est possible.
- Le dépôt d'un fichier hors de la zone est neutralisé pour éviter l'ouverture par le navigateur.
- Afficher une confirmation verte après chargement réussi (en dehors des sections repliables).
- Afficher une erreur rouge en cas de fichier non valide.

### 14.1 Sections repliables

**📋 Diagnostic de chargement** :
- Nom du fichier
- Chemin / dossier
- Lignes conservées
- Colonnes
- Période détectée

**📐 Paramètres statiques** :
- Formule cote → volume du lac de Saint-Cassien
- Débits réservés
- Seuils d'alerte
- Capacités théoriques des usines
- Disponibilités réglementaires
- Lignes de référence du lac

**🗂️ Mapping des colonnes par catégorie** :
- Colonnes classées par catégorie du cahier des charges (Ressources naturelles, Prélèvements, Capacité de production, Volumes produits, etc.)

### 14.2 Validation du fichier

- **≥ 40 colonnes mappées** et **≥ 1000 lignes de données**.
- En cas d'échec : message d'erreur de format, état vide sur toutes les pages.

## 15. Règles générales de rendu

- Police générale proche de Streamlit : sans-serif, style moderne.
- Fond clair, menu latéral gris très pâle (#f0f2f6).
- Page active avec fond légèrement bleuté (#e3ecf7).
- Cartes KPI blanches avec bordure et ombre légère.
- Formats de date : dd/mm/yyyy.
- Volumes avec séparateur de milliers français.
- Légendes Plotly interactives.
- Axes lisibles, titres de graphes identiques aux captures.

## 16. Fichier livrable

Le fichier `index.html` est autonome et contient :

- HTML pour les 6 pages
- CSS pour le style (Streamlit-like)
- JavaScript pour toute la logique :
  - Chargement Excel (SheetJS)
  - Nettoyage des données
  - Calculs métier
  - Filtre de période global
  - Rendu des graphiques (Plotly.js)
  - Persistance (IndexedDB + localStorage)
  - Validation du fichier
  - Gestion des erreurs

**Bibliothèques externes chargées via CDN :**
- SheetJS : `https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js`
- Plotly.js : `https://cdn.plot.ly/plotly-2.35.2.min.js`

**Logo :** `https://sicasil.com/wp-content/uploads/2020/09/logo-sicasil-2.jpg`