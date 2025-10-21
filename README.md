# Tirelire_Japonaise_VBA (somme des cases colorées)

Description
------------
Ce petit script Google Apps Script permet de calculer automatiquement la somme des cellules colorées dans une plage d'un Google Sheet. Il est pensé pour une "tirelire japonaise" : un tableau où chaque case colorée représente une somme à épargner. Le script additionne les valeurs numériques des cases qui ont une couleur de fond différente (autre que les couleurs neutres) et écrit le total dans une cellule dédiée.

But
----
Permettre d'épargner de manière ludique : en coloriant les cases du tableau (par exemple, une case par jour, chaque couleur représentant un montant), le total est recalculé automatiquement à chaque modification de la feuille et affiché dans une cellule de synthèse.

Fonctionnalités principales
----------------------------
- Parcourt une plage définie (par défaut `D7:S22`).
- Détecte les cellules ayant une couleur de fond différente des couleurs "neutres" (blanc / transparent / noir dans le code).
- Additionne les valeurs numériques présentes dans ces cellules.
- Écrit la somme dans une cellule de résultat (par défaut `O24`).
- S'exécute automatiquement à chaque modification grâce à la fonction `onEdit(e)`.
Installation et utilisation
---------------------------
1. Ouvrez votre Google Sheet.
2. Menu Extensions → Apps Script.
3. Créez un nouveau script et collez le code ci‑dessus.
4. Sauvegardez. Lors du premier déclenchement automatique, Google demandera d'autoriser le script à accéder à la feuille : acceptez les autorisations.
5. Dans votre feuille, utilisez la plage `D7:S22` (ou modifiez la variable `rangeNotation` dans le code) pour votre tableau de cases colorées.
6. Le total s'affichera automatiquement dans la cellule `O24`. Vous pouvez changer cette cellule dans le code si nécessaire.

Configuration (à adapter)
-------------------------
- Plage à analyser : modifier `rangeNotation = "D7:S22"` dans `onEdit`.
- Cellule résultat : modifier `sheet.getRange("O24").setValue(sum)`.
- Couleurs ignorées : le script ignore actuellement `#ffffff`, `#000000`, `#ffffff00` et `#00000000`. Ajustez cette liste si votre fond de feuille est différent ou si vous voulez ignorer d'autres couleurs.
- Format des valeurs : le script convertit le contenu de chaque cellule en nombre via `parseFloat`. Assurez‑vous que les valeurs sont bien des nombres (ou des textes convertibles).

Limites et conseils
--------------------
- Codes hex : les codes de couleur renvoyés par `getBackgrounds()` sont des hex (ex. `#ffffff`). Selon le thème ou le format, il peut y avoir des variantes (ex. `#ffffff00` pour la transparence). Adaptez la condition de filtrage si nécessaire.
- Merged cells (cellules fusionnées) : peuvent compliquer l'indexation ; testez sur votre tableau.
- Valeurs non numériques : les cellules contenant du texte non convertible seront ignorées (parseFloat retourne NaN).
- Performance : pour de très grandes plages, l'exécution à chaque modification peut être lourde. On peut optimiser en ne recalculant que quand l'édition touche la plage observée.
- Sécurité & permissions : l'exécution demande des autorisations Google. Ne partagez pas le script si vous n'êtes pas à l'aise avec les accès.
- Transparence visuelle : choisissez des couleurs claires pour les cases et des couleurs neutres pour les fonds afin d'éviter des erreurs de détection.

Améliorations possibles
-----------------------
- Faire en sorte que le script ne recalcule que si l'édition intervient dans la plage observée (vérifier `e.range`).
- Ajouter des montants par couleur (par exemple une couleur = 1€ une autre = 5€) au lieu de lire la valeur numérique dans la case.
- Afficher un petit rapport (nombre de cases coloriées).

Titre court à afficher dans ton dépôt
------------------------------------
- "Tirelire japonaise (somme des cases colorées)"
ou comme vous le souhaitez

Crédit
------
Script et README fournis par NaomieBGT.
