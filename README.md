# Automatisation du Journal de Travail

Ce dépôt contient le code source d'un outil d'automatisation développé pour gérer l'envoi de mes rapports d'activités hebdomadaires.

## Technologies utilisées
* **Langage :** JavaScript (ES6)
* **Plateforme :** Google Apps Script (Cloud / Serverless)
* **Services Google :** Sheets API, GmailApp, HTML Service

## Fonctionnalités
Ce script a pour but d'éviter les oublis et de gagner du temps chaque vendredi. Voici son flux d'exécution :

1.  **Extraction des données :** Le script lit mon fichier Google Sheet (heures, activités, problèmes rencontrés) pour la semaine en cours.
2.  **Génération de rapport :** Il convertit les données en un rapport HTML propre et responsive (CSS intégré).
3.  **Validation interactive (17h00) :**
    * Le script m'envoie un e-mail de prévisualisation.
    * L'e-mail contient des boutons **Web App** ("Envoyer" ou "Annuler").
    * *Sécurité :* Les boutons sont protégés par un mot de passe interne pour éviter les validations non autorisées.
4.  **Envoi automatique (Timeout) :** Si je ne réponds pas avant 23h58 (oubli, panne de batterie...), le script envoie automatiquement la version par défaut au maître de stage pour respecter les délais.

## Pourquoi Google Apps Script ?
J'ai choisi d'héberger ce code sur le Cloud (Apps Script) plutôt qu'en local (Python/Bash) pour assurer une **haute disponibilité**. Le script s'exécute sur les serveurs de Google via des déclencheurs temporels (Triggers), ce qui garantit l'envoi du mail même si mon ordinateur personnel est éteint ou hors ligne.

## Structure
* `code.gs` : La logique principale (Triggers, Envoi de mail, Web App `doGet`).
---
*Projet réalisé dans le cadre de mon apprentissage.*
