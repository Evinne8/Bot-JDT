// Auteur : Evinne
// Description du script:  Ce script extrait automatiquement les données du journal de travail Google Sheet pour
//    générer un rapport formaté en HTML. Il intègre une application web permettant une validation interactive 
//    par e-mail (boutons Oui/Non) et assure un envoi automatique à 23h58 en cas d'absence de réponse.
// But : Automatiser cette tâche et économiser du temps dans le futur

////////////////////////////////////////// Variables //////////////////////////////////////////
var ID_FICHIER = PropertiesService.getScriptProperties().getProperty("FILE_ID");              // ID du fichier google sheet avec les tableaux d'heures de travail
var MY_EMAIL = PropertiesService.getScriptProperties().getProperty("MY_EMAIL");               // Mon adresse Email de l'école pour avoir une possiblité de validation double
var BOSS_EMAIL = PropertiesService.getScriptProperties().getProperty("BOSS_EMAIL");           // Adresse Email des formateurs
var TEACHER_EMAIL = PropertiesService.getScriptProperties().getProperty("TEACHER_EMAIL");     // Adresse Email de mon maître de stage
var MY_PRO_EMAIL = PropertiesService.getScriptProperties().getProperty("MY_PRO_EMAIL");
var FULL_NAME = PropertiesService.getScriptProperties().getProperty("FULL_NAME");

// Mot de passe de sécurité sur les boutons. L'url de mon script est accéssible par tout le monde, donc rajouter un mot de passe dans l'url renforce la sécurité
var PASSWORD = PropertiesService.getScriptProperties().getProperty("SECRET_KEY");

// URL pour déclancher le script.
var URL_SCRIPT = PropertiesService.getScriptProperties().getProperty("SCRIPT_URL");

/**
 * Email de validation de l'envoie du rapport des heures et des activités de la semaine au professeur Xavier Carrel et au formateurs de l'EPFL
 * Cette email est envoyé à mon adresse mail de l'école et celle de l'EPFL pour que si j'ai un problème avec l'une 
 * je puisse quand même valider ou non.
 */
function AskValidation() {
  var htmlReport = GenerateHtmlReport();
  if (!htmlReport) {
    console.log("complétement vide")
    return; 

  }
  PropertiesService.getScriptProperties().setProperty("SEND_STATUS", "WAITING");

  // On ajoute le mot de passe dans le lien des boutons
  var yesURL = URL_SCRIPT + "?action=send&mdp=" + encodeURIComponent(PASSWORD);
  var noURL = URL_SCRIPT + "?action=cancel&mdp=" + encodeURIComponent(PASSWORD);

  var htmlButtons = "<br><hr><br>" +
    "<div style='text-align: center; font-family: sans-serif;'>" +
    "<p><strong>Ce mail doit-il être envoyé au patron ?</strong></p>" +
    
    // Conteneur des boutons
    "<div style='width: 100%;'>" +
      // BOUTON OUI
      "<a href='" + yesURL + "' style='display: inline-block; background-color: #28a745; color: white; padding: 14px 25px; text-decoration: none; border-radius: 5px; font-size: 16px; margin: 10px; white-space: nowrap;'>OUI, ENVOYER</a>" +
      
      // BOUTON NON
      "<a href='" + noURL + "' style='display: inline-block; background-color: #dc3545; color: white; padding: 14px 25px; text-decoration: none; border-radius: 5px; font-size: 16px; margin: 10px; white-space: nowrap;'>NON, ANNULER</a>" +
    "</div>" +

    "<p style='color: gray; font-size: 12px; margin-top: 20px;'>Sans réponse, envoi auto à 23h58.</p>" +
    "</div>";

  MailApp.sendEmail({
    to: MY_EMAIL,
    bcc: MY_PRO_EMAIL,
    subject: "ACTION REQUISE : Validation Rapport Hebdo",
    htmlBody: htmlReport + htmlButtons
  });
}

/**
 * Enclenché automatiquement entre 23h et 00h
 * Fonction qui envoie l'email si je n'ai pas confirmer son envoie (par exemple si j'ai oublié de confirmé, ou que mon téléphone n'a plus de batterie).
 */ 
function VerifyAndSend() {
  var status = PropertiesService.getScriptProperties().getProperty("SEND_STATUS");
  if (status === "WAITING") {
    SendToBoss();
  }
}

/**
 * Fonction qui se déclenche lorsque je clique sur un des bouton du mail de confirmation
 * Il vérifie le mot de passe au cas ou quelqu'un touverai l'URL du script (une sécurité en plus).
 */
function doGet(e) {
  // VERIFICATION DE SECURITE
  var mdpRecu = e.parameter.mdp;
  if (mdpRecu !== PASSWORD) {
    return HtmlService.createHtmlOutput("<h1 style='color:red; text-align:center;'>ACCÈS REFUSÉ : Mot de passe incorrect.</h1>");
  }

  var action = e.parameter.action;
  var result = "";
  
  if (action === "send") {
    var status = PropertiesService.getScriptProperties().getProperty("SEND_STATUS");
    if (status === "SEND") {
      result = "Le mail a DÉJÀ été envoyé.";
    } else {
      SendToBoss();
      result = "C'est fait ! Le mail a été envoyé.";
    }
  } else if (action === "cancel") {
    PropertiesService.getScriptProperties().setProperty("SEND_STATUS", "CANCEL");
    result = "Envoi annulé. Le bot ne fera rien.";
  }
  
  return HtmlService.createHtmlOutput("<h1 style='text-align:center; font-family:sans-serif; margin-top:50px;'>" + result + "</h1>");
}

/**
 * Fonction qui envoie le mail au Maître de stage et au formateur-fsd
 */
function SendToBoss() {
  var htmlReport = GenerateHtmlReport();
  var subject = "Journal de travail - " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yy");
  
  MailApp.sendEmail({
    to: BOSS_EMAIL,
    cc : TEACHER_EMAIL,
    bcc: MY_PRO_EMAIL, 
    subject: subject,
    htmlBody: htmlReport
  });
  PropertiesService.getScriptProperties().setProperty("SEND_STATUS", "SEND");
}

/**
 * Fonction qui génère le html du rapport en mettant le tableau des heures 
 * puis le journal de travail des journées et les problèmes si il y en a eu
 */
function GenerateHtmlReport() {
  var timezone = Session.getScriptTimeZone();
  var today = new Date();
  var day = today.getDay(); 
  var diff = today.getDate() - day + (day == 0 ? -6 : 1);
  var mondayDate = new Date(today.setDate(diff));
  var dateSheet = Utilities.formatDate(mondayDate, timezone, "dd.MM.yyyy");

  var file = SpreadsheetApp.openById(ID_FICHIER);
  var sheets = file.getSheets();
  var activeSheet = null;

  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().indexOf(dateSheet) > -1) {
      activeSheet = sheets[i];
      break;
    }
  }
  if (!activeSheet) return null;

  // On récupère les valeurs TEXTE et les valeurs RICHES ---
  var range = activeSheet.getRange("A2:K20");
  var data = range.getDisplayValues();          // Pour le tableau (texte simple)
  var richTextData = range.getRichTextValues(); // Pour les liens dans les activités !
  
  // --- STYLE ---
  var mainFont = "font-family: Arial, Helvetica, sans-serif;";
  var containerStyle = "width: 100%; max-width: 900px; margin: 0 auto; " + mainFont + " color: #333; line-height: 1.6;";
  var h2Style = "color: #2c3e50; margin-bottom: 5px;";
  var linkStyle = "color: #1967d2; text-decoration: none;";

  // --- DEBUT DU MAIL ---
  var htmlBody = "<div style='" + containerStyle + "'>";
  
  // En-tête simple
  htmlBody += "<div>";
  htmlBody += "<h2 style='" + h2Style + "'>Journal de travail</h2>";
  htmlBody += "<p style='margin-top: 0; color: #666;'>Semaine du " + dateSheet + "</p>";
  htmlBody += "<p>Bonjour,<br><br>Voici mon relevé d'heures et le résumé de mes activités.</p>";
  htmlBody += "<p>📄 <a href='https://docs.google.com/spreadsheets/d/" + ID_FICHIER + "/edit?usp=sharing' style='" + linkStyle + "'>Accéder au fichier Google Sheet</a></p>";
  htmlBody += "</div><br>";

  // --- LE TABLEAU ---
  var tableStyle = "width: 100%; border-collapse: collapse; border: 1px solid #ccc; font-size: 12px; text-align: center;";
  var thStyle = "background-color: #76a5af; color: white; padding: 10px; border: 1px solid #ccc; font-weight: bold;";
  var tdStyle = "padding: 8px; border: 1px solid #ccc; white-space: nowrap;";
  var bilanStyle = "background-color: #4a86e8; color: white; font-weight: bold; font-size: 14px; border: 1px solid #ccc; padding: 10px;";

  htmlBody += "<div style='overflow-x: auto;'>";
  htmlBody += "<table border='1' cellpadding='0' cellspacing='0' style='" + tableStyle + "'>";
  
  htmlBody += "<thead><tr>";
  htmlBody += "<th style='" + thStyle + "'>Date</th>";
  htmlBody += "<th style='" + thStyle + "'>Début</th>";
  htmlBody += "<th style='" + thStyle + "'>D. Pause</th>";
  htmlBody += "<th style='" + thStyle + "'>F. Pause</th>";
  htmlBody += "<th style='" + thStyle + "'>Fin</th>";
  htmlBody += "<th style='" + thStyle + "'>H. Sup</th>";
  htmlBody += "<th style='" + thStyle + "'>H. Jour</th>";
  htmlBody += "<th style='" + thStyle + "'>Tot. Sem.</th>";
  htmlBody += "<th style='" + thStyle + "'>Sup. Sem.</th>";
  htmlBody += "</tr></thead><tbody>";
  
  var bilanLine = null;
  for (var i = 0; i < data.length; i++) {
    var line = data[i];
    
    // Correction Bug Vendredi
    if (line[0].toString().toLowerCase().indexOf("bilan") > -1) { 
      bilanLine = line; 
      continue; 
    }

    if (line[0] !== "") { 
      var trStyle = (i % 2 === 0) ? "background-color: #f9f9f9;" : "background-color: #ffffff;";
      htmlBody += "<tr style='" + trStyle + "'>";
      htmlBody += "<td style='" + tdStyle + " text-align:left; font-weight:bold; color:#444;'>" + line[0] + "</td>";
      htmlBody += "<td style='" + tdStyle + "'>" + line[1] + "</td>"; 
      htmlBody += "<td style='" + tdStyle + "'>" + line[2] + "</td>"; 
      htmlBody += "<td style='" + tdStyle + "'>" + line[3] + "</td>"; 
      htmlBody += "<td style='" + tdStyle + "'>" + line[4] + "</td>"; 
      htmlBody += "<td style='" + tdStyle + "'>" + line[5] + "</td>"; 
      htmlBody += "<td style='" + tdStyle + "'>" + line[6] + "</td>"; 
      htmlBody += "<td style='" + tdStyle + "'>" + line[7] + "</td>"; 
      htmlBody += "<td style='" + tdStyle + "'>" + line[8] + "</td>"; 
      htmlBody += "</tr>";
    }
  }

  if (bilanLine) {
     htmlBody += "<tr>";
     htmlBody += "<td colspan='7' style='" + bilanStyle + " text-align: right;'>Bilan :</td>";
     htmlBody += "<td style='" + bilanStyle + "'>" + bilanLine[7] + "</td>";
     htmlBody += "<td style='" + bilanStyle + "'>" + bilanLine[8] + "</td>";
     htmlBody += "</tr>";
  }
  htmlBody += "</tbody></table></div>"; 
  
  // --- DETAILS ACTIVITES (AVEC LIENS) ---
  htmlBody += "<br><h3 style='color: #4a86e8; border-bottom: 1px solid #ccc; padding-bottom: 5px; margin-top: 30px;'>📝 Détails des activités</h3>";
  
  for (var i = 0; i < data.length; i++) {
    var line = data[i];
    // --- MODIFICATION 2 : On charge la ligne Riche ---
    var richLine = richTextData[i]; 

    if (!line[0] || (line[9] === "" && line[10] === "") || line[0].toString().toLowerCase().indexOf("bilan") > -1) continue;
    
    var titre = line[0].charAt(0).toUpperCase() + line[0].slice(1);
    
    // Bloc simple pour chaque jour
    htmlBody += "<div style='margin-bottom: 25px;'>";
    
    // Titre Journée
    htmlBody += "<strong style='font-size: 16px; color: #2c3e50; display: block; margin-bottom: 5px; margin-top: 15px;'>" + titre + "</strong>";
    
    // Texte Activité (Converti en HTML avec liens)
    if (line[9] !== "") {
      var texteAvecLiens = ConvertRichText(richLine[9]);
      htmlBody += "<div style='color: #333; text-align: justify;'>" + texteAvecLiens + "</div>";
    }
    
    // Alerte Problème (Converti en HTML avec liens)
    if (line[10] !== "") {
       var problemesAvecLiens = ConvertRichText(richLine[10]);
       htmlBody += "<div style='margin-top: 10px; color: #c0392b;'>";
       htmlBody += "<strong>Problème :</strong> " + problemesAvecLiens;
       htmlBody += "</div>";
    }
    htmlBody += "</div>"; 
  }

  // Pied de page
  htmlBody += "<br><hr style='border: 0; border-top: 1px solid #eee;'><br>";
  htmlBody += "Meilleures salutations,<br><strong>"+ FULL_NAME +"</strong>";
  htmlBody += "</div>"; // Fin global

  return htmlBody;
}

/**
 * Convertit le contenu d'une cellule Google Sheet (Rich Text) en HTML valide
 * Conserve les liens hypertextes et les retours à la ligne.
 */
function ConvertRichText(richTextCell) {
  if (!richTextCell) return "";
  
  var runs = richTextCell.getRuns();
  var html = "";
  
  for (var i = 0; i < runs.length; i++) {
    var run = runs[i];
    var text = run.getText();
    var url = run.getLinkUrl();
    
    // Transforme les retours à la ligne du Sheet (\n) en balises HTML (<br>)
    text = text.replace(/\n/g, '<br>');
    
    // Si lien détecté, on crée une balise <a>
    if (url) {
      html += "<a href='" + url + "' style='color: #1967d2; text-decoration: underline;' target='_blank'>" + text + "</a>";
    } else {
      // Sinon texte normal
      html += text;
    }
  }
  return html;
}
