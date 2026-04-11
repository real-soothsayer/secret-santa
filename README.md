# 🇫🇷 Road Trip 5 Jours en France — Attractions 4 étoiles et plus

Générateur de fichier Excel premium pour un road trip 5 jours en France.  
**Uniquement des attractions notées 4 étoiles et plus sur Google Maps.**  
Les Gorges du Verdon ont été retirées (déjà visitées). De nouvelles étapes majestueuses ont été ajoutées.

---

## 🗺️ Itinéraire (16 étapes sur 5 jours)

| # | Jour | Lieu | Région | Note | Catégorie |
|---|------|------|--------|------|-----------|
| 1 | JOUR 1 | Etretat — Falaises d'Aval | Normandie | ⭐ 4.7 | Falaises et sentier côtier |
| 2 | JOUR 1 | Le Mont-Saint-Michel | Normandie | ⭐ 4.7 | Monument UNESCO |
| 3 | JOUR 2 | Château de Chenonceau | Val de Loire | ⭐ 4.7 | Château Renaissance |
| 4 | JOUR 2 | Château de Chambord | Val de Loire | ⭐ 4.6 | Château Renaissance |
| 5 | JOUR 2 | Puy du Fou | Pays de la Loire | ⭐ 4.8 | Meilleur parc du monde |
| 6 | JOUR 3 | La Roque-Gageac | Périgord Noir — Dordogne | ⭐ 4.6 | Plus Beau Village de France |
| 7 | JOUR 3 | Beynac-et-Cazenac | Périgord Noir — Dordogne | ⭐ 4.7 | Château médiéval panoramique |
| 8 | JOUR 3 | Rocamadour | Occitanie — Quercy | ⭐ 4.6 | Village perché et Sanctuaire |
| 9 | JOUR 3 | Gouffre de Padirac | Occitanie — Quercy | ⭐ 4.6 | Grotte souterraine |
| 10 | JOUR 4 | Sentier des Ocres de Roussillon | Provence — Luberon | ⭐ 4.7 | Randonnée paysage martien |
| 11 | JOUR 4 | Les Baux-de-Provence | Provence — Alpilles | ⭐ 4.6 | Village perché sur rocher |
| 12 | JOUR 4 | Carrières de Lumières | Provence — Alpilles | ⭐ 4.7 | Art immersif monumental |
| 13 | JOUR 4 | Calanques de Cassis | Provence — Côte d'Azur | ⭐ 4.7 | Mer turquoise et Nature |
| 14 | JOUR 5 | Lac d'Annecy | Alpes — Haute-Savoie | ⭐ 4.8 | Lac le plus pur d'Europe |
| 15 | JOUR 5 | Cirque du Fer à Cheval | Alpes — Haute-Savoie | ⭐ 4.8 | Montagne et 30 cascades |
| 16 | JOUR 5 | Aiguille du Midi — Chamonix | Alpes — Massif du Mont-Blanc | ⭐ 4.8 | Téléphérique 3842m et Vue 360° |

**Total : ~2 700 km — ~32h de route**

---

## 🚀 Utilisation

### Prérequis

\`\`\`bash
pip install -r requirements.txt
\`\`\`

### Générer le fichier Excel

\`\`\`bash
python generate_roadtrip.py
\`\`\`

Cela crée `Road_Trip_5J_France.xlsx` avec 4 onglets :

| Onglet | Contenu |
|--------|---------|
| 🗺️ Itinéraire | Tableau des 16 étapes avec notes, photos, conseils et liens Maps |
| 🚗 Trajet | Détail des 16 trajets avec distances et durées |
| 📅 Planning | Planning journalier matin / après-midi / soir |
| ✅ Checklist | Checklist de préparation du voyage |

---

## 📋 Critères de sélection

- ✅ Uniquement des attractions **notées 4 étoiles et plus** sur Google Maps
- ✅ Circuit logique en boucle depuis Paris sur **5 jours**
- ✅ Mix paysages, patrimoine, gastronomie et bien-être
- ❌ Gorges du Verdon exclues (déjà visitées)
