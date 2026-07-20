# 📋 Roadmap — Trakt Smart Lists

Améliorations **légères** qui ne ralentissent pas l'app et ne la rendent pas confuse.
Règle d'or : l'app doit rester RAPIDE et SIMPLE. Pas de fonctionnalité qui ajoute des appels API à chaque rendu, pas de sous-menus à 3 niveaux.

---

## ✅ Phase 1 — Petits ajouts prioritaires — FAITS

- ✅ **1. 🔔 Widget "Sorties cette semaine"** sur le Dashboard (3 max, 0 appel API)
- ✅ **2. 🎲 Roulette "Je ne sais pas quoi regarder"** (pioche score ≥ 70 **dans les filtres actifs**)
- ✅ **3. 🔍 Barre de recherche** sur "Que regarder ?" et "Calendrier" (mémoire, instantané)

## ✅ Phase 2 — Idées sympas — FAITES

- ✅ **4. 🖼️ Image Wrapped partageable (PNG 1080×1350)** — Pillow, 1 génération au clic
  - 🔧 Polices DejaVu **embarquées dans `fonts/`** → accents (é, è, à…) garantis partout
- ✅ **5. 📊 Métrique "plus vieux contenu de la watchlist"** (ajouté il y a X mois)
- ✅ **6. ⭐ "Mes coups de cœur"** V2 : **TES notes perso Trakt 9-10** d'abord, communauté en secours
- ✅ **7. 🃏 BONUS "Mes bizarreries"** : écart entre TA note et la note publique (≥2)

## 🔧 Correctifs techniques — FAITS

- ✅ **▶️ Lecture en cours (méthode TPPM `expires_at`)** : % et minutes corrects même sans heartbeat
- ✅ **Ordre dashboard** : lecture → état du nettoyage → découvertes
- ✅ **BUG doublon "Que regarder ?"** : un même contenu présent dans 2 listes n'apparaît
  plus qu'**1 fois** (version à l'ajout le plus récent) — plus de double affichage avec 2 scores
- ✅ **Score enrichi par TES notes perso** : +8 si le genre est noté haut par toi, −6 si noté bas

## ✅ Phase 3 — Demande du 20/07 — FAITE

- ✅ **2. ▶️ "Tu peux finir ça ce soir"** : fantômes triés par temps restant
  (`progress × runtime`) en haut de la page Fantômes — 0 appel
- ✅ **3. 🔁 Rewatch radar** : films vus 1× il y a 3+ ans, notés ≥ 8 (dashboard) — 0 appel
- ✅ **4. 📈 Mini digest hebdo** : "Cette semaine : X ép., Y films (9h30)" sous les KPI — 0 appel
- ✅ **5. ⭐ Coups de cœur v2** : cf. Phase 2
- ✅ **1. 🗓️ Calendrier perso `/calendars/my/shows`** : **version opt-in** sur la page Calendrier —
  UN seul appel, UNIQUEMENT si tu ouvres le bloc et cliques "Charger", puis caché.
  Aucun coût au chargement normal → conforme au critère "ne pas ralentir"

---

## 📦 Phase 4 — Finalisation publique (à faire quand l'app est finie et sans bug)

- ⏳ **README pro sur GitHub** : captures d'écran des pages, liste des fonctionnalités,
  instructions de déploiement, badge, note d'attribution Trakt. Objectif : vitrine pro.
- ⏳ **Tuto pour la communauté des Alkodiques** : article de partage s'appuyant sur le
  README + captures d'écran. Ton : ouvert, jovial, convivial, axé aide et facilité
  d'utilisation. Rédaction quasi complète par l'assistant, relecture par toi.

## 💭 Idées en réflexion (non planifiées)

- 🇬🇧 **i18n FR/EN** : sélecteur de langue. Coût élevé (chaque texte à doubler dans un
  dictionnaire, ~600 chaînes), bénéfice faible (99 % d'utilisateurs FR) → au frigo,
  possible un jour si la commu s'internationalise
- 📦 **Découpage du fichier unique en modules** (`pages/`, `api.py`, `ui.py`) : aucun gain
  de vitesse (l'import ne se produit qu'au démarrage), ce serait uniquement pour la
  lisibilité. À faire seulement si le fichier devient ingérable (risque de régression sinon)

---

## ❌ Ce qu'on NE FAIT PAS (pour garder l'app légère)

- ❌ **Pages en plus** : pas d'API Explorer, pas de page Paramètres compliquée
- ❌ **Statistiques par acteur/réalisateur** : trop d'appels API TMDB (milliers), ralentit tout
- ❌ **Synchronisation automatique / cron** : nécessite un serveur hors Streamlit Cloud
- ❌ **Widget Android natif** : hors scope web. L'image PNG Wrapped suffit pour le partage
- ❌ **Métrique "% watchlist vue"** : contradictoire avec le but de l'app (nettoyer les listes)
- ❌ **Trop de filtres / sous-filtres** : on garde les filtres simples existants
- ❌ **Auto-refresh trop fréquent** : rafraîchissement au clic / à la visite de page uniquement
