// Auteur : Evinne
// Description du script:  Ce script permet d'automatiser la création de mes fiches 
// de journal de travail qui demandait de faire certaines modifications.
// But : Automatiser cette tâche et économiser du temps dans le futur
function configureNewSheet(e) {
  if (e && e.changeType === 'INSERT_GRID') {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = spreadsheet.getSheets();
    
    // Chercher l'onglet qui viens d'être créer
    var emptySheets = sheets.filter(function(f) {
      return f.getLastRow() === 0 && f.getLastColumn() === 0;
    });

    var newSheet = null;
    // S'il n'y a qu'une feuille vide, c'est forcément la bonne
    if (emptySheets.length === 1) {
      newSheet = emptySheets[0];
    } else if (emptySheets.length > 1) {
      // S'il y en a plus d'une , je prends celle active ou la dernière
      var activeSheet = spreadsheet.getActiveSheet();
      if (activeSheet.getLastRow() === 0 && activeSheet.getLastColumn() === 0) {
        newSheet = activeSheet;
      } else {
        newSheet = emptySheets[emptySheets.length - 1];
      }
    } else {
      return;
    }

    // Trouver la semaine précédente et extraire sa date
    var regex = /^S(\d+),\s*(\d{2})\.(\d{2})\.(\d{4})$/; // Format du nommage des feuilles : S1, 02.02.2026 
    var maxS = 0;
    var lastDate = null;
    var oldSheetName = null;

    for (var i = 0; i < sheets.length; i++) {
      var match = sheets[i].getName().match(regex);
      if (match) {
        var numS = parseInt(match[1], 10);
        if (numS > maxS) {
          maxS = numS;
          lastDate = new Date(match[4], parseInt(match[3], 10) - 1, match[2]);
          oldSheetName = sheets[i].getName();
        }
      }
    }

    var sheetName = "Nouvelle_Semaine";
    var nextDate = new Date();

    // Si une ancienne feuille a été trouvé, calculer la date de la semaine d'après
    if (maxS > 0 && lastDate) {
      var nextS = maxS + 1;
      nextDate = new Date(lastDate);
      nextDate.setDate(nextDate.getDate() + 7);
      
      var day = ("0" + nextDate.getDate()).slice(-2);
      var month = ("0" + (nextDate.getMonth() + 1)).slice(-2);
      var year = nextDate.getFullYear();
      
      sheetName = "S" + nextS + ", " + day + "." + month + "." + year;
    } else {
      // Sinon, prendre le lundi de cette semaine par défaut
      var currentDay = nextDate.getDay() || 7; 
      nextDate.setDate(nextDate.getDate() - currentDay + 1);
    }

    newSheet.setName(sheetName);

    // Construire les textes, les en-têtes et injecter les formules dans les cellules
    var headers = ["Date", "Début de journée", "début de la pause", "fin de la pause", "fin de journée", "heures sup", "Heures de la journée", "Total Semaine", "heures sup de la semaine", "Tâches effectué / Notes", "Problèmes"];
    newSheet.getRange(1, 1, 1, 11).setValues([headers]);

    var dayNames = ["lundi", "mardi", "mercredi", "jeudi", "vendredi"];
    for (var d = 0; d < 5; d++) {
      var dayDate = new Date(nextDate);
      dayDate.setDate(dayDate.getDate() + d);
      var dayStr = ("0" + dayDate.getDate()).slice(-2);
      var monthStr = ("0" + (dayDate.getMonth() + 1)).slice(-2);
      var yearStr = dayDate.getFullYear().toString().slice(-2); 
      var dateText = dayNames[d] + " " + dayStr + "/" + monthStr + "/" + yearStr;
      
      var row = d + 2;
      newSheet.getRange(row, 1).setValue(dateText);
      
      // Pour mardi, remplir direct les cours de maturité
      if (d === 1) {
        newSheet.getRange(row, 2).setValue("08:00");
        newSheet.getRange(row, 3).setValue("13:10");
        newSheet.getRange(row, 4).setValue("14:10");
        newSheet.getRange(row, 5).setValue("17:20");
        newSheet.getRange(row, 10).setValue("Cours Maturité");
      }

      newSheet.getRange(row, 6).setFormula('=G' + row + '-TIME(8;20;0)');
      newSheet.getRange(row, 7).setFormula('=IFERROR(E' + row + '-B' + row + '-(D' + row + '-C' + row + '); "")');
      newSheet.getRange(row, 8).setFormula(d === 0 ? '=G2' : '=G' + row + '+H' + (row - 1));
      newSheet.getRange(row, 9).setFormula(d === 0 ? '=F2' : '=F' + row + '+I' + (row - 1));
    }

    // Je m'occupe du bilan de la semaine
    newSheet.getRange(7, 1).setValue("Bilan de la semaine:");
    newSheet.getRange(7, 8).setFormula('=SUM(G2:G6)');
    newSheet.getRange(7, 9).setFormula('=H7-(TIME(8;20;0)*5)');
    newSheet.getRange(7, 10).setValue("Cette semaine était cool");

    // Faire le cumul des heures sup en liant avec l'onglet d'avant si je l'ai trouvé
    newSheet.getRange(8, 1).setValue("Cumul des heures sup :");
    var cumulFormula = oldSheetName ? '=I7+\'' + oldSheetName + '\'!I8' : '=I7';
    newSheet.getRange(8, 9).setFormula(cumulFormula);

    // Faire la mise en forme (tailles, couleurs, alignements)
    var wholeSheet = newSheet.getRange("A1:K8");
    wholeSheet.setFontWeight("normal");
    wholeSheet.setFontSize(11);
    
    newSheet.getRange("A1:K1").setBackground("#6aa84f");
    newSheet.getRange("A7:K7").setBackground("#6d9eeb");
    newSheet.getRange("A8:K8").setBackground("#f7ee93");

    newSheet.getRange("A1:K1").setVerticalAlignment("bottom").setHorizontalAlignment("left");
    newSheet.getRange("A2:I6").setVerticalAlignment("bottom").setHorizontalAlignment("right");
    newSheet.getRange("J2:K6").setVerticalAlignment("top").setHorizontalAlignment("left");

    newSheet.getRange("A7:G7").merge().setHorizontalAlignment("right").setFontSize(15).setVerticalAlignment("middle");
    newSheet.getRange("H7:I7").setHorizontalAlignment("center").setFontSize(15).setVerticalAlignment("middle");
    newSheet.getRange("J7").setHorizontalAlignment("left").setVerticalAlignment("top");
    
    newSheet.getRange("A8:H8").merge().setHorizontalAlignment("right").setFontSize(15).setVerticalAlignment("middle");
    newSheet.getRange("I8").setHorizontalAlignment("center").setFontSize(15).setVerticalAlignment("middle");

    // Formater les heures proprement
    newSheet.getRange("B2:E6").setNumberFormat("hh:mm");
    newSheet.getRange("F2:I7").setNumberFormat("[hh]:mm");
    newSheet.getRange("I8").setNumberFormat("+[hh]:mm; [hh]:mm");
    
    // Ajuster les tailles
    newSheet.setRowHeight(1, 39);            
    newSheet.setRowHeights(2, 5, 95);        
    newSheet.setRowHeights(7, 2, 43);        

    newSheet.setColumnWidth(1, 151);   
    newSheet.setColumnWidth(2, 134);   
    newSheet.setColumnWidth(3, 140);   
    newSheet.setColumnWidth(4, 116);   
    newSheet.setColumnWidth(5, 108);   
    newSheet.setColumnWidth(6, 91);    
    newSheet.setColumnWidth(7, 162);   
    newSheet.setColumnWidth(8, 115);   
    newSheet.setColumnWidth(9, 198);   
    newSheet.setColumnWidth(10, 1740); 
    newSheet.setColumnWidth(11, 1051); 
  }
}
