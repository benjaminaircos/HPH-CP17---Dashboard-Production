// URL du webhook Power Automate
const WEBHOOK_URL = "https://defaultefdc051fe1144edaa7c5efec56da13.b1.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/0c10446ba39c47b8a97cdf0d6ebf1d60/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=-7L4tPS29G5bK6TzgiCkY8-8k9xW6pJRxBXYZY_lzh8";

// Options pour les arrêts
const optionsEquipement = [
    "",
    "Robot",
    "Brosse coupelle",
    "Brosse de nettoyage",
    "Poste trémie intérieur",
    "Poste trame",
    "Poste moule",
    "Poste pince",
    "Poste tapis de déchargement",
    "Plateau tournant"
];

const optionsQualite = [
    "",
    "Défaut matière",
    "Défaut process",
    "Non-conformité",
    "Contrôle qualité",
    "Reprise",
    "Autre qualité"
];

const optionsOrganisation = [
    "",
    "Réglage machine",
    "Changement de trame",
    "Changement de teinte",
    "Changement de format",
    "Changement de teinte + format",
    "Changement d'OF",
    "Changement de lot",
    "Nettoyage / Désinfection",
    "Prise de poste",
    "Manque matière",
    "Attente équipe",
    "Pause",
    "Réunion"
];

const optionsAutres = [
    "",
    "Formation",
    "Inventaire",
    "Essai",
    "Autre"
];

// Créer une option de select
function creerOptions(options) {
    return options.map(opt => `<option value="${opt}">${opt}</option>`).join('');
}

// Générer les heures de 1h à 8h
function genererHeures() {
    const heures = [];
    for (let h = 1; h <= 8; h++) {
        heures.push(`${h}h`);
    }
    return heures;
}

// Créer les 8 lignes avec sélection d'heures (1h à 8h)
function creerLignesFixes() {
    const tableBody = document.getElementById('tableBody');
    tableBody.innerHTML = ''; // Vider le tableau
    const heures = genererHeures();
    
    for (let i = 1; i <= 8; i++) {
        const row = tableBody.insertRow();
        row.innerHTML = `
            <td>
                <select class="table-select heure" required onchange="gererDuplicationHeure()">
                    <option value="">Sélectionner...</option>
                    ${heures.map(h => `<option value="${h}">${h}</option>`).join('')}
                </select>
            </td>
            <td><input type="number" class="prod champ-quantite" min="0" placeholder=""></td>
            <td><input type="number" class="rebuts champ-quantite" min="0" placeholder=""></td>
            <td>
                <select class="table-select equipement">
                    ${creerOptions(optionsEquipement)}
                </select>
            </td>
            <td><input type="number" class="equipement_duree" min="0" placeholder=""></td>
            <td>
                <select class="table-select qualite">
                    ${creerOptions(optionsQualite)}
                </select>
            </td>
            <td><input type="number" class="qualite_duree" min="0" placeholder=""></td>
            <td>
                <select class="table-select organisation">
                    ${creerOptions(optionsOrganisation)}
                </select>
            </td>
            <td><input type="number" class="organisation_duree" min="0" placeholder=""></td>
            <td>
                <select class="table-select autres">
                    ${creerOptions(optionsAutres)}
                </select>
            </td>
            <td><input type="number" class="autres_duree" min="0" placeholder=""></td>
            <td><input type="text" class="commentaire" placeholder="Commentaire..."></td>
        `;
    }
}

// Gérer la duplication des heures (griser prod/rebuts si heure déjà utilisée)
function gererDuplicationHeure() {
    const tableBody = document.getElementById('tableBody');
    const rows = tableBody.rows;
    const heuresUtilisees = {};
    
    // Réinitialiser tous les champs
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const prodInput = row.querySelector('.prod');
        const rebutsInput = row.querySelector('.rebuts');
        
        prodInput.disabled = false;
        rebutsInput.disabled = false;
        prodInput.style.backgroundColor = '';
        rebutsInput.style.backgroundColor = '';
    }
    
    // Parcourir toutes les lignes pour identifier les doublons
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const heure = row.querySelector('.heure').value;
        
        if (heure && heure !== "") {
            if (heuresUtilisees[heure] === undefined) {
                // Première occurrence de cette heure
                heuresUtilisees[heure] = i;
            } else {
                // Heure déjà utilisée - griser prod et rebuts
                const prodInput = row.querySelector('.prod');
                const rebutsInput = row.querySelector('.rebuts');
                
                prodInput.disabled = true;
                rebutsInput.disabled = true;
                prodInput.value = '';
                rebutsInput.value = '';
                prodInput.style.backgroundColor = '#e0e0e0';
                rebutsInput.style.backgroundColor = '#e0e0e0';
            }
        }
    }
}

// Réinitialiser le formulaire
function resetForm() {
    if (confirm('Voulez-vous vraiment réinitialiser toutes les données ?')) {
        document.getElementById('date').value = '';
        document.getElementById('reference').value = '';
        document.getElementById('trigramme').value = '';
        document.getElementById('equipe').value = '';
        
        // Réinitialiser la date et l'équipe
        initialiserDate();
        determinerEquipe();
        
        // Recréer les 8 lignes fixes
        creerLignesFixes();
        
        afficherMessage('Formulaire réinitialisé', 'success');
    }
}

// Afficher un message
function afficherMessage(texte, type) {
    const messageDiv = document.getElementById('message');
    messageDiv.textContent = texte;
    messageDiv.className = `message ${type}`;
    
    if (type === 'success' || type === 'error') {
        setTimeout(() => {
            messageDiv.style.display = 'none';
        }, 5000);
    }
}

// Collecter les données du formulaire
function collecterDonnees() {
    const date = document.getElementById('date').value;
    const reference = document.getElementById('reference').value.toUpperCase().trim();
    const trigramme = document.getElementById('trigramme').value.toUpperCase().trim();
    const equipe = document.getElementById('equipe').value;
    
    // Validation des champs obligatoires
    if (!date || !reference || !trigramme || !equipe) {
        afficherMessage('⚠️ Veuillez remplir la date, la référence, le trigramme et l\'équipe', 'error');
        return null;
    }
    
    // Validation de la référence (non vide)
    if (!reference) {
        afficherMessage('⚠️ La référence est obligatoire (ex: U522003)', 'error');
        return null;
    }
    
    // Validation du trigramme (3 lettres + 1 chiffre optionnel)
    if (trigramme.length < 3 || trigramme.length > 4 || !/^[A-Z]{3}[0-9]?$/.test(trigramme)) {
        afficherMessage('⚠️ Le trigramme doit contenir 3 lettres et optionnellement 1 chiffre (ex: ABC ou ABC1)', 'error');
        return null;
    }
    
    const tableBody = document.getElementById('tableBody');
    const rows = tableBody.rows;
    
    const lignes = [];
    
    // Parcourir toutes les lignes et ne garder que celles avec une heure sélectionnée
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const heure = row.querySelector('.heure').value;
        
        // Si l'heure est sélectionnée, on ajoute la ligne
        if (heure && heure !== "") {
            lignes.push({
                heure: heure,
                prod: row.querySelector('.prod').value || "0",
                rebuts: row.querySelector('.rebuts').value || "0",
                equipement: row.querySelector('.equipement').value || "",
                equipement_duree: row.querySelector('.equipement_duree').value || "0",
                qualite: row.querySelector('.qualite').value || "",
                qualite_duree: row.querySelector('.qualite_duree').value || "0",
                organisation: row.querySelector('.organisation').value || "",
                organisation_duree: row.querySelector('.organisation_duree').value || "0",
                autres: row.querySelector('.autres').value || "",
                autres_duree: row.querySelector('.autres_duree').value || "0",
                commentaire: row.querySelector('.commentaire').value || ""
            });
        }
    }
    
    // Vérifier qu'il y a au moins une ligne à envoyer
    if (lignes.length === 0) {
        afficherMessage('⚠️ Veuillez remplir au moins une ligne avec une heure', 'error');
        return null;
    }
    
    // Récupérer les objectifs depuis le localStorage (s'ils existent)
    const objectifsStockes = JSON.parse(localStorage.getItem('objectifs_hph_cp17')) || {};
    
    return {
        date: date,
        reference: reference,
        trigramme: trigramme,
        equipe: equipe,
        objectif_prod_equipe: objectifsStockes.prodEquipe || 5000,
        objectif_taux_rebuts: objectifsStockes.tauxRebuts || 5,
        objectif_trs: objectifsStockes.trs || 85,
        cadence_instantanee: objectifsStockes.cadence || 15,
        lignes: lignes
    };
}

// Envoyer les données vers Power Automate
async function envoyerDonnees() {
    const donnees = collecterDonnees();
    
    if (!donnees) {
        return;
    }
    
    console.log('📤 Données à envoyer:', donnees);
    
    afficherMessage('⏳ Envoi des données en cours...', 'loading');
    
    try {
        const response = await fetch(WEBHOOK_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(donnees)
        });
        
        if (response.ok) {
            afficherMessage('✅ Données enregistrées avec succès dans SharePoint !', 'success');
            
            // Option : réinitialiser le formulaire après succès
            setTimeout(() => {
                if (confirm('Données enregistrées ! Voulez-vous créer une nouvelle saisie ?')) {
                    resetForm();
                }
            }, 1000);
        } else {
            throw new Error(`Erreur HTTP: ${response.status}`);
        }
    } catch (error) {
        console.error('Erreur lors de l\'envoi:', error);
        afficherMessage('❌ Erreur lors de l\'enregistrement. Vérifiez votre connexion et réessayez.', 'error');
    }
}

// Initialiser la date du jour
function initialiserDate() {
    const today = new Date();
    const dateStr = today.toISOString().split('T')[0];
    document.getElementById('date').value = dateStr;
}

// Auto-sélection de l'équipe selon l'heure
function determinerEquipe() {
    const heure = new Date().getHours();
    let equipe = '';
    
    if (heure >= 6 && heure < 14) {
        equipe = 'Matin';
    } else if (heure >= 14 && heure < 22) {
        equipe = 'Soir';
    } else {
        equipe = 'Nuit';
    }
    
    document.getElementById('equipe').value = equipe;
}

// Initialisation au chargement de la page
document.addEventListener('DOMContentLoaded', function() {
    initialiserDate();
    determinerEquipe();
    
    // Créer les 8 lignes fixes (1h à 8h)
    creerLignesFixes();
    
    console.log('✅ Application HPH CP17 initialisée');
    console.log('🔗 Connecté à Power Automate');
});

// Raccourcis clavier
document.addEventListener('keydown', function(e) {
    // Ctrl + S pour enregistrer
    if (e.ctrlKey && e.key === 's') {
        e.preventDefault();
        envoyerDonnees();
    }
});
