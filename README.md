# Système de gestion des urgences - Outil VBA sous Excel

**Simulation d’un service d’urgences hospitalières : admission du patient, triage infirmier, prise en charge médicale, gestion du personnel soignant, archivage des séjours et statistiques d’activité automatisées.**

![Excel](https://img.shields.io/badge/Excel-217346?style=for-the-badge\&logo=microsoft-excel\&logoColor=white)
![VBA](https://img.shields.io/badge/VBA-867DB1?style=for-the-badge)
![UserForms](https://img.shields.io/badge/UserForms-5C2D91?style=for-the-badge)

> Travail réalisé en équipe de quatre sous **Microsoft Excel** avec **VBA** : **feuilles de calcul** pour stocker les patients actuels, l’équipe médicale et les archives, **UserForms** pour guider chaque étape du parcours patient, **fonctions de priorisation** inspirées du triage de Manchester (P1 à P5), et **dictionnaires + graphiques générés par code** pour le module statistique.

<br>

## 🎯 Objectifs

* Reproduire le **parcours réel d’un patient** aux urgences : accueil administratif, triage infirmier, soins médicaux, sortie.
* **Automatiser la priorisation** des patients selon leur état clinique et leur temps d’attente.
* Suivre en temps réel la **présence et les horaires du personnel** soignant.
* Garantir la **traçabilité** de chaque séjour et produire des **statistiques d’activité** à jour.

<br>

## 🛠️ Compétences mobilisées

- **Programmation VBA** : procédures et fonctions, UserForms événementiels, listes déroulantes dynamiques, contrôles de saisie à la frappe, calendrier interactif pour la date de naissance.
- **Logique métier & algorithmes de tri** : `CalculerPriorite()` (âge, mode d’arrivée, motif, douleur) puis `CalculerPrioriteMedicale()` (température, tension, fréquences cardiaque et respiratoire, AVPU, douleur), avec réévaluation selon le temps d’attente ; cas particulier SMUR orienté directement en réanimation.
- **Gestion des identifiants** : attribution d’un **IPP** unique et pérenne (réutilisé si le patient existe déjà dans les archives) et d’un **IEP** distinct pour chaque séjour, afin d’éviter les doublons.
- **Gestion du personnel** : pointage arrivée/départ avec calcul automatique du temps de travail, statut Présent/Absent, ajout et suppression de membres, filtrage des soignants présents et compétents par spécialité.
- **Fiabilité des données** : blocage des incohérences (double arrivée, départ sans arrivée), validation obligatoire des champs avant enregistrement, archivage automatique des séjours terminés.
- **Reporting automatisé** : agrégation via dictionnaires (`Scripting.Dictionary`) par médecin, infirmier, mode d’arrivée, pôle et date ; génération de cinq graphiques par semaine ou par mois, à partir des périodes réellement présentes dans les archives.

<br>

## 📌 Résultats clés

L’outil couvre **l’ensemble du parcours patient sans que l’utilisateur ait à manipuler directement les feuilles Excel** : depuis la page d’accueil, chaque formulaire ouvre automatiquement le dossier du patient prioritaire, ce qui reproduit le fonctionnement d’un logiciel hospitalier. Le **double niveau de triage** (à l’admission puis après la prise des constantes vitales) permet d’affiner la priorité au fil du parcours, tandis que les patients SMUR et les retours à domicile sont traités comme dans un vrai service. Côté équipe, le pointage et la mise à jour des effectifs sont **sécurisés par des contrôles qui empêchent les saisies incohérentes**. Enfin, le module statistique est **entièrement autonome** : chaque nouveau séjour archivé est pris en compte dans les analyses suivantes sans modifier le code, ce qui donne une vision fiable de l’activité et de la charge de travail du service.
